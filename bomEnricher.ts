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

// Функция для чтения заголовков
async function readHeaders(worksheet: Worksheet): Promise<{ headers: string[], headerRow: number }> {
    // Проходим по строкам сверху вниз
    for (let i = 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        const cellValues = [];
        let hasLetters = false;
        
        // Читаем значения ячеек в строке
        for (let j = 1; j <= worksheet.columnCount; j++) {
            const cellValue = row.getCell(j).text?.trim() || '';
            cellValues.push(cellValue);
            // Проверяем наличие букв в ячейке
            if (/[a-zA-Z]/.test(cellValue)) {
                hasLetters = true;
            }
        }
        
        // Если в строке есть хотя бы одна ячейка с буквами - это заголовок
        if (hasLetters) {
            console.log(`Найдены заголовки в строке ${i}:`, cellValues);
            return {
                headers: cellValues,
                headerRow: i
            };
        }
    }
    
    throw new Error('Не удалось найти строку с заголовками');
}

// 2. Поиск информации о компоненте
async function searchComponentInfo(description: string, partNumber: string): Promise<string> {
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

USE PLAIN TEXT, no markdown!

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
   
    const response = await openai.chat.completions.create({
        model: "microsoft/phi-4",
        messages: [
            {
                role: "system",
                content: `You are an expert component description formatter. Your job is to transform raw component search results into clean, standardized descriptions following strict formatting rules. Below are the updated rules and structures you MUST follow:

        INPUT: You will receive raw search results of any length, which include component specifications and URLs. Your task is to extract relevant information and format it properly.

        OUTPUT FORMAT: You must ALWAYS return EXACTLY 3 lines:
        Line 1: Fully formatted and standardized component description (see rules below).
        Line 2: Primary source URL (must start with https://). If none available, write "NO_SOURCE".
        Line 3: Secondary source URL (must be from input, different from primary). If none available, write "NO_SECOND_SOURCE".

        === DESCRIPTION STRUCTURE ===
        The description on Line 1 must follow this structure (if applicable):
        "<TYPE> <SERIES> <VALUE> <RATED VOLTAGE> <TOLERANCE> <TEMPERATURE COEFFICIENT> <PACKAGE> <MOUNTING TYPE> <PACKAGING>"
        For RF and special components also include:
        "<DCR> <CURRENT> <Q> <SRF> <GAIN> <NOISE FIGURE> <INSERTION LOSS> <RETURN LOSS>"
        For connectors:
        "CONN <SERIES> <CONNECTOR TYPE> <FREQUENCY RANGE> <MATING MECHANISM> <PACKAGE> <MOUNTING TYPE>"
        For diodes and components with packaging details:
        "<TYPE> <SERIES> <VOLTAGE> <CURRENT> <PACKAGE> <MOUNTING TYPE> <PACKAGING>"
        For resistors:
        "RES <SERIES> <RESISTANCE> <TOLERANCE> <POWER RATING> <PACKAGE> <MOUNTING TYPE> <PACKAGING>"
        For capacitors:
        "CAP <SERIES> <CAPACITANCE> <RATED VOLTAGE> <TOLERANCE> <TEMPERATURE COEFFICIENT> <PACKAGE> <MOUNTING TYPE> <PACKAGING>"
        For inductors:
        "IND <SERIES> <INDUCTANCE> <TOLERANCE> <CURRENT> <DCR> <PACKAGE> <MOUNTING TYPE> <PACKAGING>"
        For inserts:
        "INSERT <SERIES> <THREAD SIZE> <LENGTH> <MATERIAL> <LOCKING TYPE> <FINISH>"
        For screws:
        "SCREW <THREAD SIZE> <LENGTH> <MATERIAL> <THREAD SERIES> <THREAD CLASS> <MOUNTING TYPE>"
        For unknown cases:
        - When TYPE cannot be determined, return "UNKNOWN_COMPONENT".
        - If PART NUMBER exists, include it in the output with "UNKNOWN_COMPONENT".
        
        === MANDATORY FIELDS ===
        TYPE: Always standardized based on component category:
        - Electronic Components: "CAP", "RES", "CONN", "FILTER", "IC", "XTAL", "IND", "DIPLEXER", "AMPLIFIER", "DIODE"
        - Mechanical Parts: "SCREW", "BOLT", "NUT", "WASHER", "STANDOFF", "BRACKET"
        - Chemical Materials: "ADHESIVE", "EPOXY", "PASTE", "COATING"
        - Other: "WIRE", "CABLE", "LABEL", "TAPE"

        For connectors specifically:
        - Connector Type (e.g., "SMPM", "SMA")
        - Frequency Range (e.g., "65GHZ")
        - Mating Mechanism (e.g., "PUSH-ON")
        - Mounting Type (e.g., "SMT")

        For diodes:
        - Type (e.g., "DIODE").
        - Series (if applicable, e.g., "SWTG").
        - Voltage (e.g., "80V").
        - Current (e.g., "250MA").
        - Package size (e.g., "SOT23").
        - Mounting Type (e.g., "SMT", "THT").
        - Packaging type (e.g., "T&R", "BULK").

        For resistors:
        - Type (e.g., "RES").
        - Series (if applicable, e.g., "CHIP").
        - Resistance (e.g., "0.475R").
        - Tolerance (e.g., "1%").
        - Power Rating (e.g., "1W").
        - Package size (e.g., "2512").
        - Mounting Type (e.g., "SMT").
        - Packaging type (e.g., "T&R").

        For capacitors:
        - Capacitance (e.g., "470PF").
        - Rated Voltage (e.g., "50V").
        - Tolerance (e.g., "5%").
        - Temperature Coefficient (e.g., "X7R").
        - Package size (e.g., "0402").
        - Mounting Type (e.g., "SMT").
        - Packaging type (e.g., "T&R").

        For inductors:
        - Inductance (e.g., "12NH").
        - Tolerance (e.g., "2%").
        - Current (e.g., "1.4A").
        - DCR (e.g., "93M").
        - Package size (e.g., "0402").
        - Mounting Type (e.g., "SMT").
        - Packaging type (e.g., "T&R").

        For EMI filters:
        - Type (e.g., "EMI").
        - Current (e.g., "350MA").
        - Voltage (e.g., "10V").
        - Capacitance (e.g., "250PF").
        - Package size (e.g., "0402").
        - Mounting Type (e.g., "SMT").

        For inserts:
        - Thread Size (e.g., "6-32").
        - Length (e.g., "1.5IN").
        - Material (e.g., "TYPE 304 SS").
        - Locking Type (e.g., "SELF-LOCKING").
        - Finish (if available, e.g., "NONE").

        For screws:
        - Thread Size (e.g., "6-32").
        - Length (e.g., "1.5IN").
        - Material (e.g., "TYPE 304 SS").
        - Thread Series (e.g., "UNC").
        - Thread Class (e.g., "2B").
        - Mounting Type (e.g., "SMT").


        VALUE: Format depends on component type:
        - Electronic Components: normalized (e.g., "470PF", "1.8K", "81M")
        - Mechanical Parts: dimensions/thread (e.g., "2X.250", "M3", "4-40")
        - Chemical Materials: volume/weight when specified (e.g., "50ML", "100G")
        - Inductors: Inductance value (e.g., "12NH", "10NH")
        - RF Components: Frequency range (e.g., "1.2-3GHZ")
        
        PACKAGE: Context-dependent:
        - Electronic: package size (e.g., "0402", "SOT23", "8-SMD")
        - Mechanical: head style/material (e.g., "PHIL", "HEX", "CRES")
        - Chemical: container type if applicable (e.g., "TUBE", "SYRINGE")
        - Inductors: Inductance value (e.g., "12NH", "10NH").
        - Capacitors: Capacitance value (e.g., "470PF").
        - RF Components: Frequency range (e.g., "1.2-3GHZ").


        MOUNTING TYPE:
        - Electronic: "SMT" or "THT" - REQUIRED!
        - Always include mounting type (e.g., "SMT", "THT").
        - Mechanical: mounting method if relevant (e.g., "PANEL", "CHASSIS")

        OPTIONAL FIELDS (if present):
        - SERIES: Technology or series identifier (e.g., "CRM", "COAX", "MS", "AN")
        - RATED VOLTAGE: Voltage rating (e.g., "50V", "5.5V")
        - TOLERANCE: Tolerance value (e.g., "5%", "2%")
        - TEMPERATURE COEFFICIENT: Stability indicator (e.g., "X7R", "C0G")
        - DCR: DC Resistance (e.g., "93M").
        - CURRENT: Current rating (e.g., "1.4A").
        - Q: Quality factor (e.g., "Q=30").
        - SRF: Self-resonant frequency (e.g., "SRF=5.2GHZ").
        - RATED VOLTAGE: Voltage rating (e.g., "5V").
        - GAIN: Amplifier gain (e.g., "19.8DB").
        - NOISE FIGURE: Noise figure (e.g., "NF=0.31DB").
        - INSERTION LOSS: Insertion loss (e.g., "1.2DB").
        - RETURN LOSS: Return loss (e.g., "15DB").

        === CRITICAL PROCESSING RULES ===
        1. **Preserve Meaning**: Never alter the meaning of the input. Only standardize format and terminology.
        2. **Mandatory Normalization**:
           - Join numeric values and units (e.g., "1 UF" -> "1MF").
           - Standardize resistance (e.g., "50 OHM" -> "50R").
           - Standardize capacitance (e.g., "1 uF" -> "1MF").
           - Replace temperature coefficients (e.g., "C0G" remains "C0G").
           - Frequency normalization (e.g., "1.5 GHZ" -> "1.5GHZ").
           - Remove ±/+/- from tolerances (e.g., "±2%" -> "2%").
        3. **Handle Missing Data**: If mandatory fields are missing:
           - For TYPE or VALUE: Return "NO_PART_NUMBER" only if completely unavailable.
           - For URL: Use "NO_SOURCE" or "NO_SECOND_SOURCE".
        4. **Validation of URLs**: Only use URLs from the input. Copy them exactly; do not create or modify URLs.

        === CHANGE RESISTANCE DESCRIPTION RULES (EXACT MATCHES ONLY) ===
        - "81 MOHM" -> "81M"
        - "1.8 KOHM" -> "1.8K"
        - "50 OHM" -> "50R"
        - "50 Ohm" -> "50R"
        - "50 ohm" -> "50R"
        - "50Ω", "93mΩ", "100KΩ" -> "50R", "93M", "100K"

        === CHANGE CAPACITANCE DESCRIPTION RULES (EXACT MATCHES ONLY) ===
        - "1 UF" -> "1MF"
        - "1 uF" -> "1MF"
        - "1 μF" -> "1MF"
        - "1 NF" -> "1000PF"
        - "1 nF" -> "1000PF"
        - "1 nF" -> "0.001MF"
        - "1 NF" -> "0.001MF"
        - PF stays as PF

        === ADDITIONAL OUTPUT FORMAT RULES ===
        - Replace "CER"/"CERAMIC" with "CRM"
        - Replace "nH" with "NH"
        - If ends with "SMD"/"SMT2" -> "SMT"
        - Keep component identifiers (e.g., "2W0")
        - For screws: "PHILLIPS" -> "PHIL", "STAINLESS" -> "SS"
        - For materials: "STAINLESS STEEL" -> "SS", "ALUMINUM" -> "AL"

        === UNIT STANDARDIZATION RULES (EXACT MATCHES ONLY) ===
        1. Component Type Standardization:
           - "RESIST"/"RESS"/"resist"/etc -> "RES"
           - Remove "CHIP"/"CHP"/"chip" completely
           - Replace "CER"/"CERAMIC" with "CRM"
           - "IND"/"INDUCTOR"/"CHOKE" -> "IND"
           - "CONN"/"CONNECTOR" -> "CONN"
           - "FILTER"/"FLT" -> "FILTER"
           - "CAPACITOR" -> "CAP"
           - "AMPLIFIER"/"AMP" -> "AMPLIFIER"
           - "OSCILLATOR"/"OSC" -> "OSC"
           - "TRANSFORMER"/"XFMR" -> "XFMR"

        2. Value and Unit Standardization:
           - Remove spaces between value and unit
           - Remove "±/+/-" from tolerances
           - "±/+/-2%" -> "2%"
           - "nH"/"NH" -> "NH"
           - "uH"/"μH"/"mH"/"MH" -> "MH"
           - "pF"/"PF" -> "PF"
           - "nF" -> "1000PF"
           - "uF"/"μF" -> "MF"
           - "mΩ"/"mohm"/"mOhm" -> "M"
           - "Ω"/"ohm"/"OHM" -> "R"
           - "kΩ"/"kohm"/"KOHM" -> "K"
           - "MΩ"/"Mohm"/"MOHM" -> "M"
           - "GHz"/"GHZ" -> "GHZ"
           - "MHz"/"MHZ" -> "MHZ"
           - "kHz"/"KHZ" -> "KHZ"

        3. Package and Mounting:
           - Always include mounting type ("SMT" or "THT")
           - If ends with "SMD"/"SMT2" -> "SMT"
           - Keep package identifiers (e.g., "0402", "SOT23", "2W0")
           - Include package dimensions when available
           - For filters and specific components, include impedance (e.g., "50R")

        4. Extended Component Parameters:
           - Include voltage rating when available (e.g., "50V")
           - Include current rating when available (e.g., "1.4A")
           - Include frequency range for RF components
           - Include power rating when available (e.g., "0.1W")
           - Include temperature coefficient when applicable
           - Include Q factor when available for inductors
           - Include DCR when available for inductors
           - Include resonant frequency when relevant
           - Include insertion loss and return loss for connectors
           - Include frequency range for connectors

        5. Description Structure:
           - Start with standardized component type
           - Include series/family when available
           - Include all relevant electrical parameters
           - End with package and mounting type
           - Keep manufacturer's key specifications
           - Maintain consistent order of parameters

            === HANDLING UNMENTIONED CASES ===
        1. If input data does not match predefined categories (e.g., CONNECTOR, INSERT, SCREW):
        - Extract all available key specifications and list them as "SPECIFICATIONS".
        - Return a generic description with "UNKNOWN_COMPONENT" and include all known details.

        2. If no PART NUMBER is provided:
        - Return "NO_PART_NUMBER".
        - Include all specifications in the "SPECIFICATIONS" field

        === EXAMPLES ===
        INPUT:
        RESIST CHIP 1.8 KOHM 0.06W 1% 0402 SMT
        Part: RC0402FR-071K8L
        URLs: https://digikey.com/example1, https://mouser.com/example2

        OUTPUT:
        RES 1.8K 0.06W 1% 0402 SMT
        https://digikey.com/example1
        https://mouser.com/example2

        INPUT:
        CAP CER 470 PF 50 V 5% X7R 0402 T&R
        Part: GRM155R71H471KA01D
        URLs: https://example.com

        OUTPUT:
        CAP CRM 470PF 50V 5% X7R 0402 T&R SMT
        https://example.com
        NO_SECOND_SOURCE

        INPUT:
        IC BUFFER NON-INVERTING 1.65V-5.5V 2.4NS 50PF 6PIN SOT363 SMT
        Part: BUF-IC001
        URLs: https://datasheet.example.com, https://octopart.example.com

        OUTPUT:
        IC BUFFER NON-INVERTING 1.65V-5.5V 2.4NS 50PF SOT363 SMT
        https://datasheet.example.com
        https://octopart.example.com

        INPUT:
        CHIP RESIS .475 OHM 1% 1W T&R
        Part: WSL2512R4750FTA
        URLs: https://example.com
        OUTPUT:
        RES CHIP 0.475R 1% 1W 2512 T&R SMT
        https://example.com
        NO_SECOND_SOURCE

        INPUT:
        FILTER EMI 350MA 10V 250PF 1005*
        Part: MEM1005PP251T001
        URLs: https://www.xecor.com/product/mem1005pp251t001, https://everythingrf.com/example

        OUTPUT:
        FILTER EMI 350MA 10V 250PF 1005 SMT
        https://www.xecor.com/product/mem1005pp251t001
        https://everythingrf.com/example

        INPUT:
        SCREW PH PHIL CRES NC 2X.250
        Part: MS51957-3
        URLs: https://military-fasteners.com/example

        OUTPUT:
        SCREW PHIL CRES 2X.250
        https://military-fasteners.com/example
        NO_SECOND_SOURCE

        INPUT:
        PCB ASSEMBLY GUIDE
        Part: NONE
        URLs: NONE

        OUTPUT:
        NO_PART_NUMBER
        NO_SOURCE
        NO_SECOND_SOURCE

        === RF COMPONENT EXAMPLES ===
        INPUT:
        IND 12NH ±2% 93M 0402 SMT
        Part: LQW15AN12NG8ZD
        URLs: https://www.murata.com/example

        OUTPUT:
        IND LQW15 12NH 2% 93M Q=30 SRF=5.2GHZ 0402 SMT
        https://www.murata.com/example
        NO_SECOND_SOURCE

        INPUT:
        IC AMP LNA 1.2-3GHZ TDFN 8
        Part: AHL5216T8
        URLs: https://www.asb.co.kr/example

        OUTPUT:
        IC AMPLIFIER 1.2-3GHZ 5V GAIN=19.8DB NF=0.31DB TDFN8 2.0X2.0MM
        https://www.asb.co.kr/example
        NO_SECOND_SOURCE

        INPUT:
        DIPLEXER DC-3G 1.5DB 2W 1008-8
        Part: LDPQ-132-33+
        URLs: https://www.minicircuits.com/example

        OUTPUT:
        DIPLEXER 0HZ-1.28GHZ/1.55GHZ-3GHZ 1.2DB RETURN LOSS=15DB 8-SMD SMT
        https://www.minicircuits.com/example
        NO_SECOND_SOURCE

        INPUT:
        INSERT HCOIL CRES LOCK .138-32*
        Part: NAS1130-06L15
        URLs: https://catalog.monroeaerospace.com/item/all-categories/inserts-1/nas1130-06l15-1
        OUTPUT:
        INSERT NAS1130 6-32 1.5IN TYPE 304 SS SELF-LOCKING NONE
        https://catalog.monroeaerospace.com/item/all-categories/inserts-1/nas1130-06l15-1
        https://www.minicircuits.com/example

        CRITICAL: Always validate input, preserve data meaning, and adhere to all rules strictly.`
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
        const searchResults = await searchComponentInfo(description, partNumber);
        
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

// Функция для получения списка листов из Excel файла
export async function getSheetNames(buffer: Buffer): Promise<string[]> {
    try {
        console.log('=== Начинаем получение списка листов ===');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        
        const sheets = workbook.worksheets;
        const sheetNames = sheets.map(sheet => sheet.name);
        
        // Очищаем память
        workbook.removeWorksheet(1);
        return sheetNames;
    } catch (error) {
        console.error('Ошибка при получении списка листов:', error);
        throw error;
    }
}

// Функция для чтения заголовков с конкретного листа
export async function getFileHeadersFromSheet(buffer: Buffer, sheetName: string): Promise<string[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.getWorksheet(sheetName);
    
    if (!worksheet) {
        throw new Error('Указанный лист не найден в файле');
    }

    const { headers } = await readHeaders(worksheet);
    
    // Очищаем память
    workbook.removeWorksheet(worksheet.id);
    return headers;
}

// Экспортируем функцию чтения заголовков для использования в API
export async function getFileHeaders(buffer: Buffer): Promise<string[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.getWorksheet(1);
    
    if (!worksheet) {
        throw new Error('Excel файл не содержит листов');
    }

    const { headers } = await readHeaders(worksheet);
    
    // Очищаем память
    workbook.removeWorksheet(1);
    return headers;
}

// Основная функция обработки Excel файла
export async function processExcelBuffer(
    buffer: Buffer,
    sheetName: string | undefined,
    partNumberColumn: number,
    descriptionColumn: number,
    onProgress?: (current: number, total: number) => void,
    onPreview?: (before: string, after: string, source: string) => void,
    fileId?: string
): Promise<Uint8Array> {
    console.log('Начинаем обработку Excel файла...');
    console.log('Выбранный лист:', sheetName || 'Первый лист');
    
    const workbook = new ExcelJS.Workbook();
    
    try {
        // Загружаем только выбранный лист
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(sheetName || 1);
        
        if (!worksheet) {
            throw new Error('Указанный лист не найден в файле');
        }

        // Находим строку с заголовками
        const { headers, headerRow } = await readHeaders(worksheet);

        // Проверяем валидность индексов колонок
        if (partNumberColumn < 1 || partNumberColumn > worksheet.columnCount ||
            descriptionColumn < 1 || descriptionColumn > worksheet.columnCount) {
            throw new Error('Некорректные индексы колонок');
        }

        // Создаем новый файл с одним листом
        const enrichedWorkbook = new ExcelJS.Workbook();
        const enrichedSheet = enrichedWorkbook.addWorksheet('Enriched');

        // Копируем заголовки и форматирование из исходного листа
        for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
            const sourceCol = worksheet.getColumn(colNumber);
            const targetCol = enrichedSheet.getColumn(colNumber);
            
            // Копируем точное значение заголовка
            const headerCell = worksheet.getRow(headerRow).getCell(colNumber);
            targetCol.header = headerCell.value?.toString() || '';
            
            // Копируем ширину и стиль
            targetCol.width = sourceCol.width;
            targetCol.font = { size: 9 };
        }

        // Добавляем новые колонки с минимальной шириной
        const lastColumn = worksheet.columnCount;
        const newColumns = [
            { header: 'Enriched Description', width: 30 },
            { header: 'Primary Source', width: 30 },
            { header: 'Secondary Source', width: 30 },
            { header: 'Full Search Result', width: 40 }
        ];

        // Устанавливаем форматирование для новых колонок
        newColumns.forEach((colConfig, index) => {
            const newCol = enrichedSheet.getColumn(lastColumn + index + 1);
            newCol.header = colConfig.header;
            newCol.width = colConfig.width;
            newCol.font = { size: 9 };
        });

        // Устанавливаем размер шрифта 9 для всех ячеек в первой строке (заголовки)
        const headerRowNew = enrichedSheet.getRow(1);
        headerRowNew.font = { size: 9 };

        // Копируем данные и добавляем обогащенную информацию
        const totalRows = worksheet.rowCount - headerRow;
        console.log(`Всего строк для обработки: ${totalRows}`);

        // Начинаем со следующей строки после заголовков
        const firstDataRow = headerRow + 1;
        console.log('Первая строка с данными:', firstDataRow);

        try {
            for (let rowNumber = firstDataRow; rowNumber <= worksheet.rowCount; rowNumber++) {
                const sourceRow = worksheet.getRow(rowNumber);
                const newRow = enrichedSheet.getRow(rowNumber);
                
                // Копируем существующие значения с форматированием
                for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
                    const sourceCell = sourceRow.getCell(colNumber);
                    const targetCell = newRow.getCell(colNumber);
                    targetCell.value = sourceCell.value;
                    targetCell.font = { size: 9 };
                }

                const description = sourceRow.getCell(descriptionColumn).text?.trim() || '';
                const partNumber = sourceRow.getCell(partNumberColumn).text?.trim() || '';

                if (!partNumber) {
                    console.log('Пропускаем строку с пустым Part Number');
                    onProgress?.(rowNumber - firstDataRow, totalRows);
                    continue;
                }

                const searchResult = await searchComponentInfo(description, partNumber);
                const result = await formatComponentDescription(searchResult);
                
                if (result) {
                    const newCells = [
                        { value: result.description },
                        { value: result.source },
                        { value: result.secondary_source },
                        { value: searchResult }
                    ];

                    newCells.forEach((cellConfig, index) => {
                        const cell = newRow.getCell(lastColumn + index + 1);
                        cell.value = cellConfig.value;
                        cell.font = { size: 9 };
                    });

                    onPreview?.(description, result.description, result.source);
                }

                onProgress?.(rowNumber - firstDataRow, totalRows);
            }
        } catch (error: any) {
            if (error?.status === 403 && error?.error?.message?.includes('Key limit exceeded')) {
                console.error('Превышен лимит API ключа. Сохраняем промежуточные результаты...');
                const arrayBuffer = await enrichedWorkbook.xlsx.writeBuffer();
                throw new Error('API_LIMIT_EXCEEDED:' + new Uint8Array(arrayBuffer).toString());
            }
            throw error;
        }

        console.log('Сохраняем результаты...');
        const arrayBuffer = await enrichedWorkbook.xlsx.writeBuffer();
        console.log('Обработка файла завершена успешно');
        return new Uint8Array(arrayBuffer);
        
    } catch (error) {
        if (error instanceof Error && error.message.startsWith('API_LIMIT_EXCEEDED:')) {
            return new Uint8Array(error.message.split(':')[1].split(',').map(Number));
        }
        throw error;
    }
}

