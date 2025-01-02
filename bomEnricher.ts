import ExcelJS from 'exceljs';
import OpenAI from 'openai';
import dotenv from 'dotenv';
import { StorageManager, ChatMessage } from './storage';
import WebSocket from 'ws';
import fs from 'fs';
import path from 'path';

dotenv.config();

// Функция очистки временных файлов
function cleanupTempFiles() {
    try {
        // Очищаем файлы истории чата
        const tableStorageDir = path.join(process.cwd(), 'table_storage');
        if (fs.existsSync(tableStorageDir)) {
            const files = fs.readdirSync(tableStorageDir);
            for (const file of files) {
                if (file.endsWith('.json')) {
                    fs.unlinkSync(path.join(tableStorageDir, file));
                    console.log(`Удален файл истории: ${file}`);
                }
            }
        }

        // Очищаем логи и временные файлы загрузок
        const uploadsDir = path.join(process.cwd(), 'uploads');
        if (fs.existsSync(uploadsDir)) {
            const files = fs.readdirSync(uploadsDir);
            for (const file of files) {
                fs.unlinkSync(path.join(uploadsDir, file));
                console.log(`Удален файл загрузки: ${file}`);
            }
        }

        // Создаем директории если они не существуют
        [tableStorageDir, uploadsDir].forEach(dir => {
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
                console.log(`Создана директория: ${dir}`);
            }
        });

    } catch (error) {
        console.error('Ошибка при очистке файлов:', error);
    }
}

// Экспортируем функцию очистки для использования в server.ts
export { cleanupTempFiles };

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
            model: "perplexity/llama-3.1-sonar-small-128k-online",
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
    fileId?: string
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

        const messages: OpenAIMessage[] = [
            {
                role: "system",
                content: `You are a BOM component description enricher.

FIRST RULE - MOST IMPORTANT:
Check if part number exists and is not empty.
If part number is empty or missing, ALWAYS return exactly these two lines:
NO_PART_NUMBER
NO_SOURCE

Only if part number exists, proceed with enrichment using rules below:

OUTPUT FORMAT:
Line 1: Enriched component description
Line 2: Valid source URL (must start with https://)

RULES:
1. ONLY output these 2 lines, nothing else
2. NO markdown, NO formatting
3. NO explanations or comments
4. Keep original parameters and add missing ones
5. For unknown components, keep original description
6. Source URL must be real and relevant
7. If part number is empty or missing, return exactly:
   NO_PART_NUMBER
   NO_SOURCE

UNIT CONVERSION:
- UF/uF -> MF
- NF/nF -> MF or PF  
- OHM/Ohm/ohm -> R
- KOHM/KOhm/kohm -> K
- MOHM/MOhm/mohm -> M
- CER -> CRM

Example 1 (with part number):
Part: GRM1555C1H390GA01D
Description: CAP CHIP CER 39 PF 50 V 2% COG

Response:
CAP CHIP CRM 39 PF 50 V 2% COG 0402 SMT
https://www.digikey.com/product-detail/en/GRM1555C1H390GA01D

Example 2 (no part number):
Part: 
Description: ANTENNA BASE, GROUND VER

Response:
NO_PART_NUMBER
NO_SOURCE`
            },
            ...chatHistory.map(msg => ({
                role: msg.role as 'user' | 'assistant',
                content: msg.content
            })),
            {
                role: "user", 
                content: `Enrich this component description:
Part: ${partNumber}
Description: ${description}

Return ONLY enriched description and valid source URL in 2 lines.`
            }
        ];

        const completion = await openai.chat.completions.create({
            model: "perplexity/llama-3.1-sonar-small-128k-online",
            messages
        });
        
        // Проверяем наличие ответа
        if (!completion?.choices?.[0]?.message?.content) {
            console.error('Нет ответа от LLM');
            return null;
        }
        
        const response = completion.choices[0].message.content;
        
        // Строгая валидация ответа
        const lines = response.split('\n').map(line => line.trim()).filter(line => line);
        
        if (lines.length !== 2) {
            console.error('Invalid LLM response format - expected 2 lines, got:', lines.length);
            return null;
        }

        const [description_text, source_url] = lines;

        // Проверяем, не пустой ли парт-номер
        if (description_text === 'NO_PART_NUMBER') {
            return {
                description: description,
                source: 'N/A - No part number'
            };
        }

        // Проверяем формат ответа
        if (!description_text || description_text.includes('###') || 
            !source_url || !source_url.startsWith('https://')) {
            console.error('Invalid LLM response format:', {description_text, source_url});
            return null;
        }

        // Если компонент неизвестен, возвращаем оригинальное описание
        const result = {
            description: description_text.includes('unknown') ? description : description_text,
            source: source_url
        };

        return result;

    } catch (error: any) {
        console.error('Error processing row:', error);
        throw error; // Пробрасываем ошибку выше для обработки в processExcelBuffer
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

        // Анализируем заголовки с помощью LLM
        const headerAnalysis = await openai.chat.completions.create({
            model: "perplexity/llama-3.1-sonar-small-128k-online",
            messages: [
                {
                    role: "system",
                    content: `You are analyzing the Bill of Materials table headers.
Your task is to understand which columns contain component information.

IMPORTANT: Analyze the semantic meaning of each header, not just exact matches.
Look for columns that might contain:

1. Component Description/Name:
   - Could be labeled as: Name, Description, Component, Item, etc.
   - Should contain detailed component information
   - Usually the longest text field

2. Part Number:
   - Could be labeled as: Part #, Part Number, Vendor Part, Item Number, etc.
   - Contains unique identifier for the component
   - Usually alphanumeric code

Analyze each header's meaning and purpose in the BOM context.
Return your analysis in plain text, explaining which headers you identified and why.`
                },
                {
                    role: "user",
                    content: `Analyze these table headers and explain which ones contain component descriptions and part numbers: ${headers.join(', ')}`
                }
            ]
        });

        console.log('Анализ заголовков:', headerAnalysis.choices[0].message.content);

        // Находим нужные колонки на основе анализа
        let descIndex = -1;
        let partIndex = -1;

        // Проходим по заголовкам и ищем наиболее подходящие колонки
        headers.forEach((header, index) => {
            if (!header) return; // Пропускаем пустые заголовки
            
            const headerLower = header.toString().toLowerCase();
            
            // Ищем колонку с описанием
            if (
                headerLower.includes('name') ||
                headerLower.includes('description') ||
                headerLower.includes('component') ||
                headerLower.includes('item')
            ) {
                if (descIndex === -1 || headerLower.includes('description')) {
                    descIndex = index;
                }
            }
            
            // Ищем колонку с парт-номером
            if (
                headerLower.includes('part') ||
                headerLower.includes('number') ||
                headerLower.includes('#') ||
                headerLower.includes('pn') ||
                headerLower.includes('vendor')
            ) {
                if (partIndex === -1 || headerLower.includes('part')) {
                    partIndex = index;
                }
            }
        });

        console.log(`Определены колонки: Description='${headers[descIndex]}' (${descIndex}), Part='${headers[partIndex]}' (${partIndex})`);

        if (descIndex < 0 || partIndex < 0) {
            throw new Error(`Не удалось определить нужные колонки. Найденные заголовки: ${headers.join(', ')}`);
        }

        // Добавляем колонки для результатов
        const llmSuggestionIndex = worksheet.columnCount + 1;
        const sourceIndex = worksheet.columnCount + 2;

        worksheet.getCell(1, llmSuggestionIndex).value = 'Enriched Description';
        worksheet.getCell(1, sourceIndex).value = 'Source';

        const totalRows = worksheet.rowCount - 1;
        console.log(`Всего строк для обработки: ${totalRows}`);

        try {
            // Перебираем строки, начиная со второй (пропускаем заголовки)
            for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
                const row = worksheet.getRow(rowNumber);
                
                // Получаем значения из нужных колонок
                const description = row.getCell(descIndex).value?.toString() ?? '';
                const partNumber = row.getCell(partIndex).value?.toString() ?? '';

                console.log(`\nОбработка строки ${rowNumber}/${worksheet.rowCount}:`);
                console.log(`Описание: ${description}`);
                console.log(`Парт-номер: ${partNumber}`);

                const result = await processRow(description, partNumber, fileId);
                
                if (result) {
                    console.log('Получен ответ от LLM:', result);
                    
                    // Записываем результаты в таблицу
                    row.getCell(llmSuggestionIndex).value = result.description;
                    row.getCell(sourceIndex).value = {
                        text: result.source,
                        hyperlink: result.source
                    };
                    
                    // Вызываем callback для обновления UI
                    onPreview?.(description, result.description, result.source);
                }

                onProgress?.(rowNumber - 1, totalRows);
            }
        } catch (error: any) {
            // Если превышен лимит API, сохраняем промежуточные результаты
            if (error?.status === 403 && error?.error?.message?.includes('Key limit exceeded')) {
                console.error('Превышен лимит API ключа. Сохраняем промежуточные результаты...');
                const arrayBuffer = await workbook.xlsx.writeBuffer();
                throw new Error('API_LIMIT_EXCEEDED:' + new Uint8Array(arrayBuffer).toString());
            }
            throw error;
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
                        column.width = Math.min(maxLength + 2, 100);
                    }
                }
            });
        }

        console.log('Сохраняем результаты...');
        const arrayBuffer = await workbook.xlsx.writeBuffer();
        console.log('Обработка файла завершена успешно');
        return new Uint8Array(arrayBuffer);
        
    } catch (error) {
        // Проверяем, не содержит ли ошибка промежуточные результаты
        if (error instanceof Error && error.message.startsWith('API_LIMIT_EXCEEDED:')) {
            return new Uint8Array(error.message.split(':')[1].split(',').map(Number));
        }
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
                content: `You are analyzing the "Bill of Materials" table and FILLING IN component descriptions.
                You can use information from ANY EXTERNAL websites.

                Answer user questions using all available component information.
                You can provide detailed answers, explanations, recommendations - anything that helps the user better understand the components.`
            },
            ...chatHistory.map(msg => ({
                role: msg.role as 'user' | 'assistant',
                content: msg.content
            })),
            { role: "user", content: question }
        ];

        const response = await openai.chat.completions.create({
            model: "perplexity/llama-3.1-sonar-small-128k-online",
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