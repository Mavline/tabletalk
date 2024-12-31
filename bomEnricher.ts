import ExcelJS from 'exceljs';
import OpenAI from 'openai';
import dotenv from 'dotenv';
import { StorageManager, ChatMessage } from './storage';
import WebSocket from 'ws';

dotenv.config();

const openai = new OpenAI({
    baseURL: "https://openrouter.ai/api/v1",
    apiKey: process.env.OPENROUTER_API_KEY
});

const storageManager = StorageManager.getInstance();

type OpenAIMessage = {
    role: 'system' | 'user' | 'assistant';
    content: string;
};

// Проверка подключения к API
async function checkApiConnection(): Promise<boolean> {
    try {
        const completion = await openai.chat.completions.create({
            model: "deepseek/deepseek-chat",
            messages: [
                {
                    role: "system",
                    content: "Test connection"
                },
                {
                    role: "user",
                    content: "Hello"
                }
            ]
        });
        return !!completion.choices[0]?.message?.content;
    } catch (error) {
        console.error('API Connection Error:', error);
        return false;
    }
}

// Выполняем проверку при запуске
console.log('Проверка подключения к OpenRouter API...');
checkApiConnection().then(isConnected => {
    if (isConnected) {
        console.log('✅ Подключение к OpenRouter API успешно установлено');
    } else {
        console.error('❌ Ошибка подключения к OpenRouter API');
        console.error('Проверьте OPENROUTER_API_KEY в файле .env');
    }
});

interface ProcessedRow {
    description: string;
    source: string;
}

async function processRow(
    description: string, 
    partNumber: string,
    fileId?: string,
    onLLMResponse?: (response: string, accept: () => void, reject: () => void) => void
): Promise<ProcessedRow | null> {
    try {
        // Получаем историю чата для контекста
        let chatHistory: ChatMessage[] = [];
        if (fileId) {
            try {
                chatHistory = await storageManager.getChatHistory(fileId);
            } catch (error) {
                console.log('История чата не найдена, используем пустой контекст');
            }
        }

        // Формируем сообщения с учетом истории
        const messages: OpenAIMessage[] = [
            {
                role: "system",
                content: `You are a JSON-only response API. You MUST ALWAYS respond in valid JSON format.
You are analyzing a Bill of Materials table and enriching component descriptions.

RESPONSE FORMAT:
You MUST ONLY return a JSON object with exactly two fields:
{
    "description": "formatted component description following rules below",
    "source": "full product URL with https://"
}
DO NOT include any explanations, notes or text outside the JSON object.

RULES FOR DESCRIPTIONS:
1. NEVER simplify or shorten existing descriptions
2. Keep ALL parameters from the original description
3. Add missing parameters if found in part number
4. STRICTLY FOLLOW unit conversion rules - this is MANDATORY!

UNIT CONVERSION RULES (MANDATORY!):
1. Capacitors:
   - UF/uF -> MF (microfarads to MF)
   Examples:
   * 10 UF -> 10 MF
   * 4.7 uF -> 4.7 MF
   
   - NF/nF -> Use MF or PF (shortest form)
   Examples:
   * 470 NF -> 0.47 MF
   * 1 NF -> 1000 PF
   
2. Resistors and Impedance:
   - OHM/Ohm/ohm -> R
   - KOHM/KOhm/kohm -> K
   - MOHM/MOhm/mohm -> M
   Examples:
   * 50 OHM -> 50R
   * IMPEDANCE=180 OHM -> IMPEDANCE=180R
   * 1.5 KOHM -> 1.5K
   * 2.2 MOHM -> 2.2M

3. Ceramic:
   - CER -> CRM (ALWAYS replace!)
   Examples:
   * CAP CHIP CER -> CAP CHIP CRM
   * CERAMIC -> CRM

Component Type Formats:

Capacitors:
CAP CHIP CRM <value> <power> <deviation> <temperature> <assembly> <size> <parameters>
Example: CAP CHIP CRM 39 PF 50 V 2% COG 0402 SMT Q=30

Resistors:
RES CHIP <value> <power> <deviation> <size> <assembly> <parameters>
Example: RES CHIP 1.8K 0.0625W 1% 0402 SMT

Filters:
FILTER <type> <frequency> <size> <parameters>
Example: FILTER BAND 1567.5MHZ 1DB SMT2

Inductors:
IND CHIP <value> <current> <deviation> <size> <parameters>
Example: IND CHIP 12NH 1.24A 2% 0402 Q=30

Connectors:
CONN <type> <pins> <pitch> <mount> <parameters>
Example: CONN COAX 1P 1.778MM SMT UFL 3P

RULES FOR LINKS:
1. Always provide FULL product URL with all parameters
2. Use direct product pages instead of search/category pages
3. Prefer datasheets and detailed specifications
4. Include product ID/part number in URL when possible

IMPORTANT:
1. Use ONLY information from vendor part number
2. Do NOT add manufacturer names
3. ALWAYS check and convert ALL units according to rules above
4. If original description is longer - keep ALL its parameters
5. Return strictly JSON format:
{
    "description": "formatted description following rules above",
    "source": "full product URL with https://"
}`
            },
            ...chatHistory.map(msg => ({
                role: msg.role as 'user' | 'assistant',
                content: msg.content
            })),
            {
                role: "user",
                content: `Analyze this component:
Part: ${partNumber}
Description: ${description}

Provide enriched description following ALL rules above.
IMPORTANT: Keep ALL parameters from original description!`
            }
        ];

        const completion = await openai.chat.completions.create({
            model: "deepseek/deepseek-chat",
            messages,
            response_format: { type: "json_object" }
        });

        const result = JSON.parse(completion.choices[0].message.content || '{}');

        // Добавляем https:// к ссылке, если его нет
        if (result.source && !result.source.startsWith('http')) {
            result.source = 'https://' + result.source;
        }

        if (!result.description || result.description === 'null') {
            return null;
        }

        if (onLLMResponse) {
            return new Promise((resolve) => {
                onLLMResponse(
                    result.description,
                    () => resolve(result),
                    () => resolve(null)
                );
            });
        }

        if (fileId) {
            // Сохраняем JSON для контекста
            await storageManager.addChatMessage(fileId, 'assistant', completion.choices[0].message.content || '');
        }

        return result as ProcessedRow;
    } catch (error) {
        console.error('Error processing row:', error);
        throw error;
    }
}

export async function processExcelBuffer(
    buffer: Buffer,
    onProgress?: (current: number, total: number) => void,
    onPreview?: (before: string, after: string, source: string) => void,
    fileId?: string
): Promise<Uint8Array> {
    console.log('Начинаем обработку Excel файла...');
    const workbook = new ExcelJS.Workbook();
    
    try {
        console.log('Загружаем файл в память...');
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1);
        
        if (!worksheet) {
            console.error('Не найден лист в Excel файле');
            throw new Error('Excel файл не содержит листов');
        }

        console.log('Анализируем структуру файла...');
        const headers = worksheet.getRow(1).values as string[];
        console.log('Найдены заголовки:', headers);

        // Даем LLM проанализировать заголовки
        const headerAnalysis = await openai.chat.completions.create({
            model: "openai/gpt-4o-2024-11-20",
            messages: [
                {
                    role: "system",
                    content: `Вы анализируете заголовки таблицы Bill of Materials.
Найдите колонки с описанием компонента и его парт-номером.
Ответ дайте в формате JSON:
{
    "descriptionColumn": "индекс колонки с описанием",
    "partNumberColumn": "индекс колонки с парт-номером"
}`
                },
                {
                    role: "user",
                    content: `Заголовки таблицы: ${headers.join(', ')}`
                }
            ],
            response_format: { type: "json_object" }
        });

        const headerInfo = JSON.parse(headerAnalysis.choices[0].message.content || '{}');
        const nameIndex = parseInt(headerInfo.descriptionColumn) - 1;
        const partIndex = parseInt(headerInfo.partNumberColumn) - 1;

        console.log(`Индексы колонок: Description=${nameIndex}, Part=${partIndex}`);

        if (isNaN(nameIndex) || isNaN(partIndex) || nameIndex < 0 || partIndex < 0) {
            throw new Error(`Не найдены нужные колонки. Найденные колонки: ${headers.join(', ')}`);
        }

        // Добавляем только две колонки
        const llmSuggestionIndex = worksheet.columnCount + 1;
        const sourceIndex = worksheet.columnCount + 2;

        worksheet.getCell(1, llmSuggestionIndex).value = 'Enriched Description';
        worksheet.getCell(1, sourceIndex).value = 'Source';

        const totalRows = worksheet.rowCount - 1;
        console.log(`Всего строк для обработки: ${totalRows}`);

        for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
            const row = worksheet.getRow(rowNumber);
            const description = row.getCell(nameIndex + 1).value?.toString();
            const partNumber = row.getCell(partIndex + 1).value?.toString();

            console.log(`\nОбработка строки ${rowNumber}/${worksheet.rowCount}:`);
            console.log(`Описание: ${description}`);
            console.log(`Парт-номер: ${partNumber}`);

            if (description && partNumber) {
                try {
                    console.log('Отправляем запрос к LLM...');
                    const result = await processRow(description, partNumber, fileId);
                    if (result) {
                        console.log('Получен ответ от LLM:', result);
                        row.getCell(llmSuggestionIndex).value = result.description;
                        
                        // Создаем кликабельную ссылку в Excel с полным URL
                        const sourceCell = row.getCell(sourceIndex);
                        sourceCell.value = {
                            text: result.source,
                            hyperlink: result.source
                        };
                        
                        onPreview?.(description, result.description, result.source);
                    } else {
                        console.log('LLM не предложила изменений для этой строки');
                        row.getCell(llmSuggestionIndex).value = 'No suggestions';
                        row.getCell(sourceIndex).value = '-';
                    }
                } catch (error) {
                    console.error(`Ошибка при обработке строки ${rowNumber}:`, error);
                    row.getCell(llmSuggestionIndex).value = `Ошибка: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`;
                    row.getCell(sourceIndex).value = 'ERROR';
                }
            } else {
                console.log('Пропускаем строку - пустое описание или парт-номер');
                row.getCell(llmSuggestionIndex).value = 'Пропущено - неполные данные';
                row.getCell(sourceIndex).value = '-';
            }

            onProgress?.(rowNumber - 1, totalRows);
        }

        // Автоматически подгоняем ширину колонок
        if (worksheet.columns) {
            worksheet.columns.forEach(column => {
                let maxLength = 0;
                if (column && typeof column.eachCell === 'function') {
                    column.eachCell({ includeEmpty: true }, cell => {
                        const length = cell.value ? cell.value.toString().length : 10;
                        if (length > maxLength) {
                            maxLength = length;
                        }
                    });
                    if (typeof column.width === 'number' || column.width === undefined) {
                        column.width = Math.min(maxLength + 2, 100); // Максимальная ширина 100 символов
                    }
                }
            });
        }

        console.log('Сохраняем результаты...');
        const arrayBuffer = await workbook.xlsx.writeBuffer();
        console.log('Обработка файла завершена успешно');
        return new Uint8Array(arrayBuffer);
        
    } catch (error) {
        console.error('Ошибка при обработке файла:', error);
        throw error;
    }
}

export async function askLLM(question: string, fileId?: string): Promise<string> {
    try {
        // Получаем историю чата для контекста
        let chatHistory: ChatMessage[] = [];
        if (fileId) {
            try {
                chatHistory = await storageManager.getChatHistory(fileId);
            } catch (error) {
                console.log('История чата не найдена, используем пустой контекст');
            }
        }

        const messages: OpenAIMessage[] = [
            {
                role: "system",
                content: `Вы анализируете таблицу "Bill of Materials" и ЗАПОЛНЯЕТЕ ОПИСАНИЕ компонентов.
Вы можете использовать информацию с ЛЮБЫХ ВНЕШНИХ сайтов

Отвечайте на вопросы пользователя, используя всю доступную информацию о компонентах.
Можете давать развернутые ответы, объяснения, рекомендации - всё, что поможет пользователю лучше понять компоненты.`
            },
            ...chatHistory.map(msg => ({
                role: msg.role as 'user' | 'assistant',
                content: msg.content
            })),
            { role: "user", content: question }
        ];

        const response = await openai.chat.completions.create({
            model: "openai/gpt-4o-2024-11-20",
            messages,
        });

        const answer = response.choices[0]?.message?.content || 'Нет ответа от LLM';

        // Сохраняем ответ в историю чата
        if (fileId) {
            await storageManager.addChatMessage(fileId, 'assistant', answer);
        }

        return answer;
    } catch (error) {
        throw new Error(`Ошибка при запросе к LLM: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`);
    }
} 