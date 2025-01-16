import ExcelJS, { Worksheet } from 'exceljs';
import OpenAI from 'openai';
import dotenv from 'dotenv';
import { StorageManager } from './storage';
import fs from 'fs';
import path from 'path';
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

interface ProcessedRow {
    description: string;
    source: string;
    secondary_source: string;
}

type ChatMessage = {
    role: 'system' | 'user' | 'assistant';
    content: string;
};

interface ColumnScore {
    colIndex: number;
    partNumberScore: number;
    descriptionScore: number;
}

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

// 1. Анализ структуры таблицы и извлечение данных
async function analyzeTableStructure(worksheet: Worksheet, headerRowIndex: number): Promise<{
    descIndex: number;
    partIndex: number;
    headers: string[];
    sampleRows: string[][];
}> {
    console.log('\n=== Начало анализа структуры таблицы ===');
    console.log('Чтение заголовков из строки:', headerRowIndex);
    
    const headers: string[] = [];
    const sampleRows: string[][] = [];
    let validRowsCount = 0;

    // Читаем заголовки
    const headerRow = worksheet.getRow(headerRowIndex);
    for (let col = 1; col <= worksheet.columnCount; col++) {
        const header = headerRow.getCell(col).value?.toString().trim() ?? '';
        headers.push(header);
    }
    console.log('Прочитаны заголовки:', headers);

    // Собираем примеры данных
    console.log('\nСбор примеров данных:');
    for (let i = headerRowIndex + 1; validRowsCount < 5 && i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        const rowValues = [];
        
        for (let j = 1; j <= headers.length; j++) {
            const cellValue = row.getCell(j).value;
            rowValues.push(cellValue ? String(cellValue).trim() : '');
        }
        
        if (rowValues.some(val => val !== '')) {
            sampleRows.push(rowValues);
            validRowsCount++;
            console.log(`Строка ${validRowsCount}:`, rowValues.join(' | '));
        }
    }

    // Анализируем паттерны в данных
    const columnScores = headers.map((_, colIndex) => {
        let partNumberScore = 0;
        let descriptionScore = 0;

        sampleRows.forEach(row => {
            const value = row[colIndex]?.trim() || '';
            
            // Проверяем характеристики парт-номера
            if (/^[A-Z0-9-+]+$/.test(value) && value.length > 5) {
                partNumberScore++;
            }
            
            // Проверяем характеристики описания
            if (/^[A-Z\s]+/.test(value) && value.includes(' ')) {
                descriptionScore++;
            }
        });

        return { colIndex, partNumberScore, descriptionScore };
    });

    // Находим колонки с максимальными очками
    const maxPartScore = Math.max(...columnScores.map(s => s.partNumberScore));
    const maxDescScore = Math.max(...columnScores.map(s => s.descriptionScore));

    const partIndexCol = columnScores.find(s => s.partNumberScore === maxPartScore);
    const descIndexCol = columnScores.find(s => s.descriptionScore === maxDescScore);

    if (!partIndexCol || !descIndexCol) {
        throw new Error('Не удалось определить нужные колонки на основе анализа данных');
    }

    const partIndex = partIndexCol.colIndex + 1;
    const descIndex = descIndexCol.colIndex + 1;

    console.log('\nРезультаты анализа:');
    console.log('Оценки колонок:', columnScores);
    console.log('Определены индексы:', { descIndex, partIndex });
    console.log('=== Завершение анализа структуры таблицы ===\n');
    
    return {
        descIndex,
        partIndex,
        headers,
        sampleRows
    };
}

// Функция для повторных попыток
async function retryOperation<T>(
    operation: () => Promise<T>,
    maxAttempts: number = 3,
    delay: number = 1000
): Promise<T> {
    let lastError;
    
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            return await operation();
        } catch (error: any) {
            lastError = error;
            
            // Проверяем, является ли ошибка ECONNRESET
            if (error.code === 'ECONNRESET' && attempt < maxAttempts) {
                console.log(`Попытка ${attempt} из ${maxAttempts} не удалась, повторяем через ${delay}мс...`);
                await new Promise(resolve => setTimeout(resolve, delay));
                continue;
            }
            
            throw error;
        }
    }
    
    throw lastError;
}

// 2. Поиск информации о компоненте
async function searchComponentInfo(partNumber: string, description: string): Promise<string> {
    console.log('\n=== Поиск информации о компоненте ===');
    console.log(`Поиск для: ${partNumber}`);
    console.log(`Исходное описание: ${description}`);

    try {
        const response = await retryOperation(async () => {
            return await openai.chat.completions.create({
                model: "perplexity/llama-3.1-sonar-small-128k-online",
                messages: [
                    {
                        role: "system",
                        content: `You are a component search engine. Return ONLY specifications and direct URLs.

DO NOT USE:
- References like [1], [2], etc.

CRITICAL URL RULES:
- Return ONLY REAL, WORKING URLs you find during search
- URLs MUST start with https://
- NO placeholder or example URLs
- NO shortened URLs
- NO text-only links
- If URL not found, return "NO_SOURCE"

Return ONLY:
1. Basic specifications
2. Key parameters
3. TOP the most relevant source URLs (NOT MORE, FULL URLs only)`
                    },
                    {
                        role: "user",
                        content: `Find information for:
Part Number: ${partNumber}
Description: ${description}

Return concise specs and REAL, FULL URLs only.`
                    }
                ],
                temperature: 0.1
            });
        });

        if (!response.choices || !response.choices[0] || !response.choices[0].message) {
            console.log('Получен пустой ответ от API');
            return 'NO_SOURCE';
        }

        const searchResult = response.choices[0].message.content || 'NO_SOURCE';
        console.log('\nРезультаты поиска:');
        console.log(searchResult);
        
        return searchResult;
    } catch (error) {
        console.error('Ошибка при поиске информации:', error);
        return 'NO_SOURCE';
    }
}

// 3. Форматирование и стандартизация описания
async function formatComponentDescription(searchResults: string): Promise<ProcessedRow> {
    console.log('\n=== Анализ и форматирование результатов поиска ===');
    console.log('Входные данные для анализа:', searchResults);

    const response = await openai.chat.completions.create({
        model: "meta-llama/llama-3.3-70b-instruct",
        messages: [
            {
                role: "system",
                content: `You are a component description formatter. Your task is to format search results into a standardized 3-line output for a table.

INPUT: You will receive search results of any length containing component specifications and URLs.

OUTPUT FORMAT: You must ALWAYS return exactly 3 lines:
Line 1: Formatted component description following all rules below
Line 2: Primary source URL (must start with https://)
Line 3: Secondary source URL (must start https://)

CRITICAL RULES:
1. Process input of ANY LENGTH - whether it's 2 lines or 100 lines
2. Always output EXACTLY 3 lines as specified above
3. If no valid URLs found, use "NO_SOURCE"
4. If no valid description possible, use "NO_PART_NUMBER"

Line 1: Standardized component description
Line 2: Primary source URL (must be from preferred sources)
Line 3: Secondary source URL (must be different from primary, or "NO_SECOND_SOURCE" if not found)

CRITICAL: NEVER CHANGE THE BEGINNING OF THE DESCRIPTION - ONLY IMPROVE SPECIFICATIONS ACCORDING TO THE RULES BELOW
- DON'T INCLUDE PART NUMBER IN THE DESCRIPTION!
- NEVER use +/- symbols, use % for tolerance
- "SMT" MUST always be the LAST word in description
- Remove all "~" symbols, use "-" for ranges
- Use EXACT part number from input, don't suggest alternatives
- Keep original frequency/voltage/current values, don't list ranges
- If part number exists - ALWAYS search and return description

CRITICAL URL RULES:
- Use ONLY URLs provided in the input
- NO placeholder or example URLs
- NO modified or shortened URLs
- If no URL in input, use "NO_SOURCE"
- NEVER create or modify URLs
- COPY URLs exactly as they appear in input

VALIDATION RULES:
Return concise specs and REAL, FULL URLs only, IF NOT POSSIBLE - return "NO_SOURCE"
If part number is empty -> return exactly:
"NO_PART_NUMBER"
"NO_SOURCE"
"NO_SECOND_SOURCE"

UNIT STANDARDIZATION RULES (EXACT MATCHES ONLY):
1. Component Type:
- "RESIST"/"RESS"/"resist"/etc -> "RES"
- Remove "CHIP"/"CHP"/"chip" completely
- Replace "CER"/"CERAMIC" with "CRM"
- IND, INDICATOR -> "IND"
- CONN, CONNECTOR -> "CONN"
- FILTER -> "FILTER"
- If ends with "SMD"/"SMT2" -> "SMT"
- Keep component identifiers (e.g. "2W0")

2. Join values and units without spaces:
- "1.575 GHZ" -> "1.575GHZ"
- "100 MA" -> "100MA"
- "50 V" -> "50V"

CHANGE RESISTANCE DESCRIPTION RULES (EXACT MATCHES ONLY):
* "81 MOHM" -> "81M"
* "1.8 KOHM" -> "1.8K"
* "50 OHM" -> "50R"
* "50 Ohm" -> "50R"
* "50 ohm" -> "50R"
* "50Ω", "93mΩ" -> "50R", "93M"

CHANGE CAPACITANCE DESCRIPTION RULES (EXACT MATCHES ONLY):
* "1 UF" -> "1MF"
* "1 uF" -> "1MF"
* "1 μF" -> "1MF"
* "1 NF" -> "1000PF"
* "1 nF" -> "1000PF"
* PF stays as PF

OUTPUT FORMAT:
- Replace "CER"/"CERAMIC" with "CRM"
- Replace "nH" with "NH"
- If ends with "SMD"/"SMT2" -> "SMT"
- Keep component identifiers (e.g. "2W0")

EXAMPLES:
1. Input: "RESIST CHIP 1.8 KOHM 0.06W 1% 04*"
Part: "RC0402FR-071K8L"
Output: "RES 1.8K 0.06W 1% 0402 SMT"
"https://www.digikey.com/product-detail/example1"
"https://www.mouser.com/product-detail/example2"

2. Input: "CAP CHP CER 470 PF 50 V 5% X7*"
Part: "GRM155R71H471KA01D"
Output: "CAP CRM 470PF 50V 5% X7R 0402 SMT"
"https://www.digikey.com/product-detail/example3"
"NO_SECOND_SOURCE"

3. Input: "DIPLEXER DC-3G 1.5DB 2W 1008-8"
Part: "LDPQ-132-33+"
Output: "DIPLEXER 0-3GHZ 1.5DB 2W 1008 SMT"
"https://www.digikey.com/product-detail/example4"
"https://www.mouser.com/product-detail/example5"

4. Input: "PCB ASSEMBLY INSTRUCTIONS"
Part: ""
Output:
"NO_PART_NUMBER"
"NO_SOURCE"
"NO_SECOND_SOURCE"`
            },
            {
                role: "user",
                content: searchResults
            }
        ],
        temperature: 0.1
    });

    const formattedText = response.choices[0]?.message?.content?.trim() || '';
    console.log('\nРазмышления модели форматирования:');
    console.log('1. Анализирую найденные спецификации...');
    console.log('2. Проверяю соответствие форматам единиц измерения...');
    console.log('3. Стандартизирую обозначения компонентов...');
    console.log('4. Проверяю и форматирую URL источников...');
    console.log('\nРезультат форматирования:');
    console.log(formattedText);

    const lines = formattedText.split('\n')
        .map(line => line.trim())
        .filter(line => line && !line.startsWith('```') && !line.includes('```'));
    
    console.log('Количество строк в ответе:', lines.length);
    console.log('=== Завершение форматирования ===\n');

    return {
        description: lines[0] || 'NO_PART_NUMBER',
        source: lines[1] || 'NO_SOURCE',
        secondary_source: lines[2] || 'NO_SOURCE'
    };
}

// Обновленная функция processRow
async function processRow(description: string, partNumber: string): Promise<ProcessedRow | null> {
    console.log('\n=== Начало обработки строки ===');
    console.log('Входные данные:');
    console.log('Description:', description);
    console.log('Part Number:', partNumber);

    try {
        if (!partNumber || partNumber.trim() === '') {
            console.log('Пропуск: отсутствует парт-номер');
            return {
                description: '',
                source: '',
                secondary_source: ''
            };
        }

        // 1. Поиск информации
        console.log('\nШаг 1: Поиск информации');
        const searchResults = await searchComponentInfo(partNumber, description);
        
        // 2. Форматирование результатов
        console.log('\nШаг 2: Форматирование результатов');
        const formattedResult = await formatComponentDescription(searchResults);
        
        console.log('\nИтоговый результат:');
        console.log(formattedResult);
        console.log('=== Завершение обработки строки ===\n');
        
        return formattedResult;

    } catch (error) {
        console.error('Ошибка при обработке строки:', error);
        return null;
    }
}

// Основная функция обработки Excel файла
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
        
        // Ищем начало таблицы в первых 10 строках
        let headerRowIndex = 0;
        let headers: string[] = [];
        
        for (let i = 1; i <= Math.min(10, worksheet.rowCount); i++) {
            const row = worksheet.getRow(i);
            const rowValues = row.values as string[];
            if (rowValues && rowValues.length > 0) {
                // Проверяем, похожа ли строка на заголовок (содержит ключевые слова или форматирование)
                const potentialHeaders = rowValues.filter(Boolean).map(val => String(val).trim());
                if (potentialHeaders.some(header => 
                    header.toLowerCase().includes('pn') || 
                    header.toLowerCase().includes('description') ||
                    header.toLowerCase().includes('mfg') ||
                    header.toLowerCase().includes('part'))) {
                    headerRowIndex = i;
                    headers = potentialHeaders;
                    break;
                }
            }
        }

        if (headerRowIndex === 0) {
            throw new Error('Не удалось найти строку с заголовками');
        }

        console.log('Найдена строка заголовков:', headerRowIndex);
        console.log('Заголовки:', headers);

        // Собираем примеры данных из следующих строк после заголовков
        const sampleRows: string[][] = [];
        let validRowsCount = 0;
        
        // Читаем до тех пор, пока не найдем 5 непустых строк или не дойдем до конца таблицы
        for (let i = headerRowIndex + 1; validRowsCount < 5 && i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            const rowValues = [];
            
            // Собираем все значения из строки
            for (let j = 1; j <= headers.length; j++) {
                const cellValue = row.getCell(j).value;
                rowValues.push(cellValue ? String(cellValue).trim() : '');
            }
            
            // Проверяем, что строка содержит хотя бы одно непустое значение
            if (rowValues.some(val => val !== '')) {
                sampleRows.push(rowValues);
                validRowsCount++;
                
                // Выводим для отладки
                console.log(`Sample Row ${validRowsCount}:`, rowValues.join(' | '));
            }
        }

        if (sampleRows.length === 0) {
            throw new Error('Не удалось найти данные для анализа структуры');
        }

        // Анализируем структуру с помощью LLM
        const { descIndex, partIndex } = await analyzeTableStructure(worksheet, headerRowIndex);

        console.log(`Определены колонки:`);
        console.log(`Description: '${headers[descIndex - 1]}' (${descIndex})`);
        console.log(`Part Number: '${headers[partIndex - 1]}' (${partIndex})`);

        // Проверяем валидность индексов
        if (descIndex < 1 || descIndex > 16384 || 
            partIndex < 1 || partIndex > 16384) {
            throw new Error(`Некорректные индексы колонок: desc=${descIndex}, part=${partIndex}`);
        }

        // Добавляем колонки для результатов (используем текущее количество + 1)
        const llmSuggestionIndex = worksheet.columnCount + 1;
        const sourceIndex = worksheet.columnCount + 2;
        const secondarySourceIndex = worksheet.columnCount + 3;
        const fullSearchResultIndex = worksheet.columnCount + 4;

        worksheet.getCell(1, llmSuggestionIndex).value = 'Enriched Description';
        worksheet.getCell(1, sourceIndex).value = 'Primary Source';
        worksheet.getCell(1, secondarySourceIndex).value = 'Secondary Source';
        worksheet.getCell(1, fullSearchResultIndex).value = 'Full Search Result';

        const totalRows = worksheet.rowCount - 1;
        console.log(`Всего строк для обработки: ${totalRows}`);

        try {
            for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
                const row = worksheet.getRow(rowNumber);
                
                // Правильно читаем данные из определенных колонок
                const description = row.getCell(descIndex).value?.toString().trim() ?? '';
                const partNumber = row.getCell(partIndex).value?.toString().trim() ?? '';

                if (!partNumber) {
                    onProgress?.(rowNumber - 1, totalRows);
                    continue;
                }

                // Поиск информации
                const searchResult = await searchComponentInfo(partNumber, description);
                
                // Форматирование результатов
                const result = await formatComponentDescription(searchResult);
                
                if (result) {
                    // Записываем форматированные результаты
                    row.getCell(llmSuggestionIndex).value = result.description;
                    row.getCell(sourceIndex).value = {
                        text: result.source,
                        hyperlink: result.source
                    };
                    row.getCell(secondarySourceIndex).value = {
                        text: result.secondary_source,
                        hyperlink: result.secondary_source
                    };
                    // Записываем полный результат поиска
                    row.getCell(fullSearchResultIndex).value = searchResult;
                    
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

        const response = await retryOperation(async () => {
            return await openai.chat.completions.create({
                model: "perplexity/llama-3.1-sonar-small-128k-online",
                messages: messages,
                temperature: 0.5,
            });
        });

        return response.choices[0]?.message?.content || 'No response generated';
    } catch (error) {
        console.error('Error in askLLM:', error);
        throw error;
    }
}
