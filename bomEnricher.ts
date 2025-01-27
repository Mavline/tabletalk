import ExcelJS, { Worksheet } from 'exceljs';
import OpenAI from 'openai';
import dotenv from 'dotenv';
import fs from 'fs';
import path from 'path';
import WebSocket from 'ws';

dotenv.config();

const openai = new OpenAI({
    baseURL: "https://openrouter.ai/api/v1",
    apiKey: process.env.OPENROUTER_API_KEY
});

// Расширяем класс StorageManager
class StorageManager {
    private static instance: StorageManager;
    private readonly maxStorageAge = 24 * 60 * 60 * 1000; // 24 часа
    private readonly maxFiles = 20; // Максимум 20 последних версий
    private readonly storageDir = path.join(process.cwd(), 'storage');
    private readonly resultsDir = path.join(this.storageDir, 'results');

    private constructor() {
        // Создаем директории при инициализации
        [this.storageDir, this.resultsDir].forEach(dir => {
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
            }
        });
        
        // Запускаем периодическую очистку каждый час
        setInterval(() => this.cleanupOldFiles(), 60 * 60 * 1000);
    }

    // Получение списка доступных файлов для конкретного ID
    getAvailableFiles(fileId: string): Array<{ name: string, timestamp: string, size: number }> {
        try {
            if (!fs.existsSync(this.resultsDir)) return [];

            return fs.readdirSync(this.resultsDir)
                .filter(file => file.startsWith(fileId))
                .map(file => {
                    const filePath = path.join(this.resultsDir, file);
                    const stats = fs.statSync(filePath);
                    return {
                        name: file,
                        timestamp: new Date(stats.mtimeMs).toISOString(),
                        size: stats.size
                    };
                })
                .sort((a, b) => b.timestamp.localeCompare(a.timestamp));
        } catch (error) {
            console.error('Ошибка при получении списка файлов:', error);
            return [];
        }
    }

    // Получение конкретного файла по ID и имени
    getFile(fileId: string, fileName: string): Buffer | null {
        try {
            const filePath = path.join(this.resultsDir, fileName);
            if (!fs.existsSync(filePath) || !fileName.startsWith(fileId)) {
                return null;
            }
            return fs.readFileSync(filePath);
        } catch (error) {
            console.error('Ошибка при чтении файла:', error);
            return null;
        }
    }

    saveIntermediateResult(fileId: string, data: Uint8Array): void {
        try {
            if (!fs.existsSync(this.resultsDir)) {
                fs.mkdirSync(this.resultsDir, { recursive: true });
            }
            
            // Добавляем временную метку к имени файла
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const fileName = `${fileId}_${timestamp}_intermediate.xlsx`;
            const filePath = path.join(this.resultsDir, fileName);
            
            fs.writeFileSync(filePath, Buffer.from(data));
            console.log(`Промежуточные результаты сохранены в ${fileName}`);

            // Проверяем и очищаем старые файлы
            this.cleanupOldFiles();
        } catch (error) {
            console.error('Ошибка при сохранении промежуточных результатов:', error);
        }
    }

    private cleanupOldFiles(): void {
        try {
            if (!fs.existsSync(this.resultsDir)) return;

            const files = fs.readdirSync(this.resultsDir)
                .map(file => ({
                    name: file,
                    path: path.join(this.resultsDir, file),
                    stats: fs.statSync(path.join(this.resultsDir, file))
                }))
                .sort((a, b) => b.stats.mtimeMs - a.stats.mtimeMs);

            const now = Date.now();
            const fileGroups = new Map<string, typeof files>();

            // Группируем файлы по fileId
            files.forEach(file => {
                const fileId = file.name.split('_')[0];
                if (!fileGroups.has(fileId)) {
                    fileGroups.set(fileId, []);
                }
                fileGroups.get(fileId)?.push(file);
            });

            // Обрабатываем каждую группу отдельно
            fileGroups.forEach(groupFiles => {
                groupFiles.forEach((file, index) => {
                    // Удаляем файлы старше 24 часов или если превышен лимит файлов для группы
                    if (now - file.stats.mtimeMs > this.maxStorageAge || index >= this.maxFiles) {
                        fs.unlinkSync(file.path);
                        console.log(`Удален старый файл: ${file.name}`);
                    }
                });
            });
        } catch (error) {
            console.error('Ошибка при очистке старых файлов:', error);
        }
    }

    static getInstance(): StorageManager {
        if (!StorageManager.instance) {
            StorageManager.instance = new StorageManager();
        }
        return StorageManager.instance;
    }
}

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
                model: "perplexity/llama-3.1-sonar-large-128k-online",
                messages: [
                    {
                        role: "system",
                        content: `You are a component search engine. Return ONLY specifications ACCORDING THE RULES below and direct URLs.


DO NOT USE:
- References like [1], [2], etc.
- Markdown formatting

CRITICAL URL RULES:
- Return ONLY REAL, WORKING URLs you find during search
- URLs MUST start with https://
- NO placeholder or example URLs
- NO shortened URLs
- NO text-only links
- If URL not found, return "NO_SOURCE"

Return ONLY:
1. Basic short standardized specifications of the electronic component:
    === DESCRIPTION STRUCTURE ===
            The description on Line 1 must follow this structure (if applicable):
            "<TYPE> <VALUE> <RATED VOLTAGE> <TOLERANCE> <TEMPERATURE COEFFICIENT> <PACKAGE> <MOUNTING TYPE>"
            For RF and special components also include:
            "<DCR> <CURRENT> <Q> <SRF> <GAIN> <NOISE FIGURE> <INSERTION LOSS> <RETURN LOSS>"
            For connectors:
            "CONN <CONNECTOR TYPE> <FREQUENCY RANGE> <MATING MECHANISM> <PACKAGE> <MOUNTING TYPE>"
            For diodes and components with packaging details:
            "<TYPE> <VOLTAGE> <CURRENT> <PACKAGE> <MOUNTING TYPE>"
            For resistors:
            "RES <RESISTANCE> <TOLERANCE> <POWER RATING> <PACKAGE> <MOUNTING TYPE>"
            For capacitors:
            "CAP <CATEGORY> <CAPACITANCE> <RATED VOLTAGE> <TOLERANCE> <TEMPERATURE COEFFICIENT> <PACKAGE> <MOUNTING TYPE>"
            For inductors:
            "IND <CATEGORY> <INDUCTANCE> <TOLERANCE> <CURRENT> <DCR> <PACKAGE> <MOUNTING TYPE>"
            For inserts:
            "INSERT <CATEGORY> <THREAD SIZE> <LENGTH> <MATERIAL> <LOCKING TYPE> <FINISH>"
            For screws:
            "SCREW <THREAD SIZE> <LENGTH> <MATERIAL> <THREAD SERIES> <THREAD CLASS> <MOUNTING TYPE>"
            For LEDs:
            "LED <COLOR> <WAVELENGTH/CCT> <LUMINOUS INTENSITY/FLUX> <FORWARD VOLTAGE> <FORWARD CURRENT> <VIEWING ANGLE> <PACKAGE> <MOUNTING TYPE>"
            For sensors:
            "SENSOR <TYPE> <AXIS> <SUPPLY VOLTAGE> <SENSITIVITY> <BANDWIDTH> <PACKAGE> <MOUNTING TYPE>"
            For MOSFETs:
            "MOSFET <TYPE> <CHANNEL> <VOLTAGE> <CURRENT> <RDS ON> <GATE THRESHOLD> <PACKAGE> <MOUNTING TYPE>"

    === ADDITIONAL OUTPUT FORMAT RULES ===
        - Replace "CER"/"CERAMIC" with "CRM"
        - Replace "nH" with "NH"
        - Remove special characters: "Ω", "+", "-", "±" from component descriptions
        - Remove "±" from all positions and descriptions (e.g., "±100PPM/K" -> "100PPM/K")
        - If ends with "SMD"/"SMT2" -> "SMT"
        - Keep component identifiers (e.g., "2W0")

    === MANDATORY FIELDS ===
            TYPE: Always standardized based on component category:
            - Electronic Components: "CAP", "RES", "CONN", "FILTER", "IC", "XTAL", "IND", "DIPLEXER", "AMPLIFIER", "DIODE", "LED", "SENSOR"
            - Mechanical Parts: "SCREW", "BOLT", "NUT", "WASHER", "STANDOFF", "BRACKET"
            - Chemical Materials: "ADHESIVE", "EPOXY", "PASTE", "COATING"
            - Other: "WIRE", "CABLE", "LABEL", "TAPE"

            For connectors specifically:
            - Connector Type: Specify the exact type of connector (e.g., "SMPM", "SMA", "U.FL").
            - Frequency Range: Include the frequency range if applicable (e.g., "65GHZ", "6GHZ").
            - Mating Mechanism: Describe the connection mechanism (e.g., "PUSH-ON", "SOLDER", "THREADED").
            - Impedance: Always include impedance if relevant (e.g., "50R", "75R").
            - Mounting Type: Must always end with the mounting type (e.g., "SMT", "THT").

            For diodes:
            - Type (e.g., "DIODE")
            - Series (if applicable)
            - Voltage (e.g., "80V")
            - Current (e.g., "250MA")
            - Package size (e.g., "SOT23")
            - Mounting Type (e.g., "SMT", "TH")

            For resistors:
            - Type (e.g., "RES")
            - Resistance (e.g., "0.475R")
            - Tolerance (e.g., "1%")
            - Power Rating (e.g., "1W")
            - Package size (e.g., "2512")
            - Mounting Type (e.g., "SMT", "TH")

            For capacitors:
            - Capacitance (e.g., "470PF")
            - Rated Voltage (e.g., "50V")
            - Tolerance (e.g., "5%")
            - Temperature Coefficient (e.g., "X7R", "C0G", "Y5V", "Z5U", "NP0", "X5R", "Y5U", "Z5V")
            - Package size (e.g., "0402")
            - Mounting Type (e.g., "SMT","TH")

            For inductors:
            - Inductance (e.g., "12NH")
            - Tolerance (e.g., "2%")
            - Current (e.g., "1.4A")
            - DCR (e.g., "93M")
            - Package size (e.g., "0402", "0603", "0805", "1206", "1210", "1812", "2010", "2512")
            - Mounting Type (e.g., "SMT", "TH")

            For EMI filters:
            - Type: Specify "EMI"
            - Current: Include the current rating (e.g., "350MA")
            - Voltage: Include the voltage rating (e.g., "10V")
            - Capacitance: Include the capacitance value (e.g., "250PF")
            - Tolerance: Include capacitance tolerance, but remove the **±** sign (e.g., "10%")
            - Package size: Specify the package size (e.g., "0402")
            - Mounting Type: Must always end with mounting type (e.g., "SMT")

            For inserts:
            - Thread Size (e.g., "6-32")
            - Length (e.g., "1.5IN")
            - Material (e.g., "TYPE 304 SS")
            - Locking Type (e.g., "SELF-LOCKING")
            - Finish (if available, e.g., "NONE")

            For screws:
            - Thread Size (e.g., "6-32").
            - Length (e.g., "1.5IN")
            - Material (e.g., "TYPE 304 SS")
            - Thread Series (e.g., "UNC")
            - Thread Class (e.g., "2B")
            - Mounting Type (e.g., "SMT")

            For labels:
            - Label type (e.g., "THT-72-457-PRINTABLE")
            - Description (e.g., "Thermal transfer printable polyimide label B-457 series, 1.75" x 0.25", 10000 labels per roll")
            - Source URL (e.g., "https://www.bradyid.com/labels/tht-72-457")
            
        MOUNTING TYPE:
        - Electronic: "SMT" or "THT" - REQUIRED!
        - Always include mounting type (e.g., "SMT", "THT")
        - Mechanical: mounting method if relevant (e.g., "PANEL", "CHASSIS")

        === CHANGE RESISTANCE DESCRIPTION RULES ===
        - "81 MOHM" -> "81M"
        - "1.8 KOHM" -> "1.8K"
        - "50 OHM" -> "50R"
        - "50 Ohm" -> "50R"
        - "50 ohm" -> "50R"
        - "50Ω" -> "50R"
        - "93mΩ" -> "93M"
        - "100KΩ" -> "100K"
        - "1MΩ" -> "1M"
        - "mΩ" -> "M"
        - "Ω" -> "R"
        - "kΩ" -> "K"
        - "MΩ" -> "M"

        === CHANGE CAPACITANCE DESCRIPTION RULES ===
        - "1UF" -> "1MF"      (microfarads -> megafarads)
        - "1uF" -> "1MF"
        - "1μF" -> "1MF"
        - "0.1UF" -> "0.1MF"
        - "0.01UF" -> "0.01MF"
        - "150NF" -> "0.15MF" (nanofarads -> megafarads, 150/1000 = 0.15)
        - "100NF" -> "0.1MF"  (100/1000 = 0.1)
        - "10NF" -> "0.01MF"  (10/1000 = 0.01)
        - "1NF" -> "1000PF"   (1 * 1000 = 1000)
        - "0.1NF" -> "100PF"  (0.1 * 1000 = 100)
        - "0.01NF" -> "10PF"  (0.01 * 1000 = 10)
        - PF stays as PF (uppercase only)

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
           - "nF"/"NF" -> convert to "PF" or "MF" based on value
           - "uF"/"μF"/"UF" -> "MF"
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

        === TEMPERATURE COEFFICIENT AND VOLTAGE/CURRENT RANGES ===
            - Temperature Coefficient format:
            - Remove ± from PPM values (e.g., "±100PPM/K" -> "100PPM/K")
            - Always keep "/K" or "/°C" suffix
            - Examples: "100PPM/K", "200PPM/°C"

            - Voltage Range format:
            - Remove ± from voltage ranges (e.g., "±5V-±50V" -> "5V-50V")
            - For dual ranges use "/" (e.g., "5V-50V/10V-100V")
            - Examples: "5V-50V", "3.3V-24V"

            - Current format:
            - Remove ± from current values (e.g., "±50MA" -> "50MA")
            - Always use uppercase units (MA, A)
            - Examples: "50MA", "1.5A"

            DO NOT confuse these with tolerance values (which use % symbol).                
2. TOP the most relevant source URLs (NOT MORE, FULL URLs only)
`
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
async function formatComponentDescription(searchResults: string, partNumber: string, description: string): Promise<ProcessedRow> {
    console.log('\n=== Анализ и форматирование результатов поиска ===');
   
    try {
        const response = await retryOperation(async () => {
            return await openai.chat.completions.create({
        model: "deepseek/deepseek-r1-distill-llama-70b",
        messages: [
            {
                role: "system",
                        content: `You are an expert component description formatter. Based on the original {description}, enhance it with information from {searchResults}, following the rules below. Format the output into EXACTLY 3 lines:
Line 1: Enhanced component description following the standardized format
Line 2: Primary source URL (must start with https://)
Line 3: Secondary source URL (must be different from primary, or NO_SECOND_SOURCE)

        INPUT: You will receive:
        1. Original component description that should be used as the base
        2. Raw search results of any length with additional specifications and URLs
        Your task is to enhance the original description with relevant information in standardized format from search results.

DO NOT Include URLs in the component description
DO NOT Include "Series" in the component description
DO NOT Include ${partNumber} in the component description
DO NOT Output any markdown formatting
DO NOT Output any additional lines
DO NOT Repeat URLs in the output
DO NOT Include any explanatory text


        OUTPUT FORMAT: You must ALWAYS return EXACTLY 3 lines:
        Line 1: Fully formatted and standardized component description (see rules below).
        Line 2: Primary source URL (must start with https://). If none available, write "NO_SOURCE".
        Line 3: Secondary source URL (must be from input, different from primary). If none available, write "NO_SECOND_SOURCE".

        === DESCRIPTION STRUCTURE ===
        The description on Line 1 must follow this structure (if applicable):
        "<TYPE> <VALUE> <RATED VOLTAGE> <TOLERANCE> <TEMPERATURE COEFFICIENT> <PACKAGE> <MOUNTING TYPE>"
        For RF and special components also include:
        "<DCR> <CURRENT> <Q> <SRF> <GAIN> <NOISE FIGURE> <INSERTION LOSS> <RETURN LOSS>"
        For connectors:
        "CONN <CONNECTOR TYPE> <FREQUENCY RANGE> <MATING MECHANISM> <PACKAGE> <MOUNTING TYPE>"
        For diodes and components with packaging details:
        "<TYPE> <VOLTAGE> <CURRENT> <PACKAGE> <MOUNTING TYPE>"
        For resistors:
        "RES <RESISTANCE> <TOLERANCE> <POWER RATING> <PACKAGE> <MOUNTING TYPE>"
        For capacitors:
        "CAP <CAPACITANCE> <RATED VOLTAGE> <TOLERANCE> <TEMPERATURE COEFFICIENT> <PACKAGE> <MOUNTING TYPE>"
        For inductors:
        "IND <INDUCTANCE> <TOLERANCE> <CURRENT> <DCR> <PACKAGE> <MOUNTING TYPE>"
        For inserts:
        "INSERT <THREAD SIZE> <LENGTH> <MATERIAL> <LOCKING TYPE> <FINISH>"
        For screws:
        "SCREW <THREAD SIZE> <LENGTH> <MATERIAL> <THREAD SERIES> <THREAD CLASS> <MOUNTING TYPE>"
        For LEDs:
        "LED <COLOR> <WAVELENGTH/CCT> <LUMINOUS INTENSITY/FLUX> <FORWARD VOLTAGE> <FORWARD CURRENT> <VIEWING ANGLE> <PACKAGE> <MOUNTING TYPE>"
        For sensors:
        "SENSOR <TYPE> <AXIS> <SUPPLY VOLTAGE> <SENSITIVITY> <BANDWIDTH> <PACKAGE> <MOUNTING TYPE>"
        For MOSFETs:
        "MOSFET <CHANNEL> <VOLTAGE> <CURRENT> <RDS ON> <GATE THRESHOLD> <PACKAGE> <MOUNTING TYPE>"
        For unknown cases:
        - When TYPE cannot be determined, return "UNKNOWN_COMPONENT".
        - If PART NUMBER exists, include it in the output with "UNKNOWN_COMPONENT".
        
        === MANDATORY FIELDS ===
        TYPE: Always standardized based on component category:
        - Electronic Components: "CAP", "RES", "CONN", "FILTER", "IC", "XTAL", "IND", "DIPLEXER", "AMPLIFIER", "DIODE", "LED", "SENSOR"
        - Mechanical Parts: "SCREW", "BOLT", "NUT", "WASHER", "STANDOFF", "BRACKET"
        - Chemical Materials: "ADHESIVE", "EPOXY", "PASTE", "COATING"
        - Other: "WIRE", "CABLE", "LABEL", "TAPE"

        For connectors specifically:
        - Connector Type: Specify the exact type of connector (e.g., "SMPM", "SMA", "U.FL")
        - Frequency Range: Include the frequency range if applicable (e.g., "65GHZ", "6GHZ")
        - Mating Mechanism: Describe the connection mechanism (e.g., "PUSH-ON", "SOLDER", "THREADED")
        - Impedance: Always include impedance if relevant (e.g., "50R", "75R")
        - Mounting Type: Must always end with the mounting type (e.g., "SMT", "THT")

        For diodes:
        - Type (e.g., "DIODE")
        - Series (if applicable)
        - Voltage (e.g., "80V")
        - Current (e.g., "250MA")
        - Package size (e.g., "SOT23")
        - Mounting Type (e.g., "SMT", "TH")

        For resistors:
        - Type (e.g., "RES")
        - Resistance (e.g., "0.475R")
        - Tolerance (e.g., "1%")
        - Power Rating (e.g., "1W")
        - Package size (e.g., "2512")
        - Mounting Type (e.g., "SMT", "TH")

        For capacitors:
        - Capacitance (e.g., "470PF")
        - Rated Voltage (e.g., "50V")
        - Tolerance (e.g., "5%")
        - Temperature Coefficient (e.g., "X7R", "C0G", "Y5V", "Z5U", "NP0", "X5R", "Y5U", "Z5V")
        - Package size (e.g., "0402")
        - Mounting Type (e.g., "SMT","TH")

        For inductors:
        - Inductance (e.g., "12NH")
        - Tolerance (e.g., "2%")
        - Current (e.g., "1.4A")
        - DCR (e.g., "93M")
        - Package size (e.g., "0402", "0603", "0805", "1206", "1210", "1812", "2010", "2512")
        - Mounting Type (e.g., "SMT", "TH")

        For EMI filters:
        - Type: Specify "EMI"
        - Current: Include the current rating (e.g., "350MA")
        - Voltage: Include the voltage rating (e.g., "10V")
        - Capacitance: Include the capacitance value (e.g., "250PF")
        - Tolerance: Include capacitance tolerance, but remove the **±** sign (e.g., "10%")
        - Package size: Specify the package size (e.g., "0402")
        - Mounting Type: Must always end with mounting type (e.g., "SMT")

        For inserts:
        - Thread Size (e.g., "6-32")
        - Length (e.g., "1.5IN")
        - Material (e.g., "TYPE 304 SS")
        - Locking Type (e.g., "SELF-LOCKING")
        - Finish (if available, e.g., "NONE")

        For screws:
        - Thread Size (e.g., "6-32")
        - Length (e.g., "1.5IN")
        - Material (e.g., "TYPE 304 SS")
        - Thread Series (e.g., "UNC")
        - Thread Class (e.g., "2B")
        - Mounting Type (e.g., "SMT")

        For labels:
        - Label type (e.g., "THT-72-457-PRINTABLE")
        - Description (e.g., "Thermal transfer printable polyimide label B-457 series, 1.75" x 0.25", 10000 labels per roll")
        - Source URL (e.g., "https://www.bradyid.com/labels/tht-72-457")

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
        - Inductors: Inductance value (e.g., "12NH", "10NH")
        - Capacitors: Capacitance value (e.g., "470PF")
        - RF Components: Frequency range (e.g., "1.2-3GHZ")

        MOUNTING TYPE:
        - Electronic: "SMT" or "THT" - REQUIRED!
        - Always include mounting type (e.g., "SMT", "THT")
        - Mechanical: mounting method if relevant (e.g., "PANEL", "CHASSIS")

        OPTIONAL FIELDS (if present):
        - SERIES: Technology or series identifier (e.g., "CRM", "COAX", "MS", "AN")
        - RATED VOLTAGE: Voltage rating (e.g., "50V", "5.5V")
        - TOLERANCE: Tolerance value (e.g., "5%", "2%")
        - TEMPERATURE COEFFICIENT: Stability indicator (e.g., "X7R", "C0G")
        - DCR: DC Resistance (e.g., "93M")
        - CURRENT: Current rating (e.g., "1.4A")
        - Q: Quality factor (e.g., "Q=30")
        - SRF: Self-resonant frequency (e.g., "SRF=5.2GHZ")
        - RATED VOLTAGE: Voltage rating (e.g., "5V")
        - GAIN: Amplifier gain (e.g., "19.8DB")
        - NOISE FIGURE: Noise figure (e.g., "NF=0.31DB")
        - INSERTION LOSS: Insertion loss (e.g., "1.2DB")
        - RETURN LOSS: Return loss (e.g., "15DB")

        === CRITICAL PROCESSING RULES ===
        1. Preserve Meaning: Never alter the meaning of the input. Only standardize format and terminology
        2. Mandatory Normalization:
           - Join numeric values and units (e.g., "1 UF" -> "1MF")
           - Standardize resistance (e.g., "50 OHM" -> "50R")
           - Standardize capacitance (e.g., "1 uF" -> "1MF")
           - Replace temperature coefficients (e.g., "C0G" remains "C0G")
           - Frequency normalization (e.g., "1.5 GHZ" -> "1.5GHZ")
           - Remove ±/+/- from tolerances (e.g., "±2%" -> "2%")
           - For Tolerance in EMI filters:
           - Remove the ± sign from everywhere (e.g., "±10%" → "10%", "±5%" → "5%", "±1%" → "1%")

        3. Handle Missing Data: If mandatory fields are missing:
           - For TYPE or VALUE: Return "NO_PART_NUMBER" only if completely unavailable
           - For URL: Use "NO_SOURCE" or "NO_SECOND_SOURCE"
        4. Validation of URLs: Only use URLs from the input. Copy them exactly; do not create or modify URLs

        === CHANGE RESISTANCE DESCRIPTION RULES ===
        - "81 MOHM" -> "81M"
        - "1.8 KOHM" -> "1.8K"
        - "50 OHM" -> "50R"
        - "50 Ohm" -> "50R"
        - "50 ohm" -> "50R"
        - "50Ω" -> "50R"
        - "93mΩ" -> "93M"
        - "100KΩ" -> "100K"
        - "1MΩ" -> "1M"
        - "mΩ" -> "M"
        - "Ω" -> "R"
        - "kΩ" -> "K"
        - "MΩ" -> "M"

        === CHANGE CAPACITANCE DESCRIPTION RULES ===
        - "1UF" -> "1MF"      (microfarads -> megafarads)
        - "1uF" -> "1MF"
        - "1μF" -> "1MF"
        - "0.1UF" -> "0.1MF"
        - "0.01UF" -> "0.01MF"
        - "150NF" -> "0.15MF" (nanofarads -> megafarads, 150/1000 = 0.15)
        - "100NF" -> "0.1MF"  (100/1000 = 0.1)
        - "10NF" -> "0.01MF"  (10/1000 = 0.01)
        - "1NF" -> "1000PF"   (1 * 1000 = 1000)
        - "0.1NF" -> "100PF"  (0.1 * 1000 = 100)
        - "0.01NF" -> "10PF"  (0.01 * 1000 = 10)
        - PF stays as PF (uppercase only)

        === ADDITIONAL OUTPUT FORMAT RULES ===
        - Replace "CER"/"CERAMIC" with "CRM"
        - Replace "nH" with "NH"
        - Remove special characters: "Ω", "+", "-", "±" from component descriptions
        - Remove "±" from all positions and descriptions (e.g., "±100PPM/K" -> "100PPM/K")
        - If ends with "SMD"/"SMT2" -> "SMT"
        - Keep component identifiers (e.g., "2W0")

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
           - "nF"/"NF" -> convert to "PF" or "MF" based on value
           - "uF"/"μF"/"UF" -> "MF"
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

        === TEMPERATURE COEFFICIENT AND VOLTAGE/CURRENT RANGES ===
            - Temperature Coefficient format:
            - Remove ± from PPM values (e.g., "±100PPM/K" -> "100PPM/K")
            - Always keep "/K" or "/°C" suffix
            - Examples: "100PPM/K", "200PPM/°C"

            - Voltage Range format:
            - Remove ± from voltage ranges (e.g., "±5V-±50V" -> "5V-50V")
            - For dual ranges use "/" (e.g., "5V-50V/10V-100V")
            - Examples: "5V-50V", "3.3V-24V"

            - Current format:
            - Remove ± from current values (e.g., "±50MA" -> "50MA")
            - Always use uppercase units (MA, A)
            - Examples: "50MA", "1.5A"

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
        CAP CRM 470PF 50V 5% X7R 0402 SMT
        https://example.com
        NO_SECOND_SOURCE

        INPUT:
        IC BUFFER NON-INVERTING 1.65V-5.5V 2.4NS 50PF 6PIN SOT363 SMT
        Part: BUF-IC001
        URLs: https://datasheet.example.com, https://octopart.example.com
        OUTPUT:
        IC BUFFER NON-INVERTING 1.65V-5.5V 2.4NS 50PF SOT363-6 SMT
        https://datasheet.example.com
        https://octopart.example.com

        INPUT:
        CHIP RESIS .475 OHM 1% 1W T&R
        Part: WSL2512R4750FTA
        URLs: https://example.com     
        OUTPUT:
        RES CHIP 0.475R 1% 1W 2512 SMT
        https://example.com
        NO_SECOND_SOURCE

        INPUT:
        FILTER EMI 350MA 10V 250PF ±10% 0402 SMT
        Part: MEM1005PP251T001
        URLs: https://www.xecor.com/product/mem1005pp251t001
        OUTPUT:
        FILTER EMI 350MA 10V 250PF 10% 0402 SMT
        https://www.xecor.com/product/mem1005pp251t001
        NO_SECOND_SOURCE


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

        === CONN EXAMPLES ===
        INPUT:
        CONN RECT 10 S SMT 1.27MM
        Part: SFM-105-02-S-D-K-TR
        URLs: https://www.samtec.com/products/sfm-105-02-s-d-k-tr
        OUTPUT:
        CONN SFM FEMALE 10POS 1.27MM DUAL ROW GOLD SMT
        https://www.samtec.com/products/sfm-105-02-s-d-k-tr
        NO_SECOND_SOURCE

        INPUT:
        IC AMP LNA 1.2-3GHZ TDFN 8
        Part: AHL5216T8
        URLs: https://www.asb.co.kr/example
        OUTPUT:
        IC AMPLIFIER 1.2-3GHZ 5V GAIN=19.8DB NF=0.31DB TDFN8 2.0X2.0MM
        https://www.asb.co.kr/example
        NO_SECOND_SOURCE

        === LABEL EXAMPLES ===
        INPUT:
        PRINTABLE LABEL,THERMAL TRANSFER,POLYMID
        Part: THT-72-457
        URLs: https://www.bradyid.com/labels/tht-72-457
        OUTPUT:
        LABEL 1.75INX0.25IN POLYIMIDE WHITE 10000ROLL labels per roll
        https://www.bradyid.com/labels/tht-72-457
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

        CRITICAL: Always validate input, preserve data meaning, and adhere to all rules strictly.

        === PACKAGE FORMAT RULES FOR IC COMPONENTS ===
        For IC components only:
        - Convert "24-PIN SOP" -> "SOP-24"
        - Convert "16-TSSOP" -> "TSSOP-16"
        - Convert "8-SOT" -> "SOT-8"
        - Convert "16-SOIC" -> "SOIC-16"
        - Convert "24-QSOP" -> "QSOP-24"
        - Convert "16-TTSOT" -> "SOT-16"
        - Always move pin count to the end of package name
        - Format: <PACKAGE_TYPE>-<PIN_COUNT>

        === ADDITIONAL OUTPUT FORMAT RULES ===
        - Replace "CER"/"CERAMIC" with "CRM"
        - Replace "nH" with "NH"
        - Remove special characters: "Ω", "+", "-", "±" from component descriptions
        - Remove "±" from all positions and descriptions (e.g., "±100PPM/K" -> "100PPM/K")
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
           - "nF"/"NF" -> convert to "PF" or "MF" based on value
           - "uF"/"μF"/"UF" -> "MF"
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
        CAP CRM 470PF 50V 5% X7R 0402 SMT
        https://example.com
        NO_SECOND_SOURCE

        INPUT:
        IC BUFFER NON-INVERTING 1.65V-5.5V 2.4NS 50PF 6PIN SOT363 SMT
        Part: BUF-IC001
        URLs: https://datasheet.example.com, https://octopart.example.com

        OUTPUT:
        IC BUFFER NON-INVERTING 1.65V-5.5V 2.4NS 50PF SOT363-6 SMT
        https://datasheet.example.com
        https://octopart.example.com

        INPUT:
        CHIP RESIS .475 OHM 1% 1W T&R
        Part: WSL2512R4750FTA
        URLs: https://example.com
        
        OUTPUT:
        RES CHIP 0.475R 1% 1W 2512 SMT
        https://example.com
        NO_SECOND_SOURCE

        INPUT:
        FILTER EMI 350MA 10V 250PF ±10% 0402 SMT
        Part: MEM1005PP251T001
        URLs: https://www.xecor.com/product/mem1005pp251t001

        OUTPUT:
        FILTER EMI 350MA 10V 250PF 10% 0402 SMT
        https://www.xecor.com/product/mem1005pp251t001
        NO_SECOND_SOURCE


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

        === CONNECTOR EXAMPLES ===
        INPUT:
        CONN RECT 10 S SMT 1.27MM
        Part: SFM-105-02-S-D-K-TR
        URLs: https://www.samtec.com/products/sfm-105-02-s-d-k-tr
        OUTPUT:
        CONN SFM FEMALE 10POS 1.27MM DUAL ROW GOLD SMT
        https://www.samtec.com/products/sfm-105-02-s-d-k-tr
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
                content: `Original description: ${description}:
Search results to format: ${searchResults}`
            }
        ],
        temperature: 0.1
    });
        }, 3, 1000); // 3 попытки с задержкой 2 секунды

    const formattedText = response.choices[0]?.message?.content?.trim() || '';
    console.log('\nРазмышления модели форматирования:');
    console.log('1. Анализирую найденные спецификации...');
    console.log('2. Проверяю соответствие форматам единиц измерения...');
    console.log('3. Стандартизирую обозначения компонентов...');
    console.log('4. Проверяю и форматирую URL источников...');

    console.log(formattedText);

    const lines = formattedText.split('\n')
        .map(line => line.trim())
        .filter((line, index) => {
            if (index === 0 && line.startsWith('**') && line.endsWith('**')) {
                return false; // Удаляем только первую строку если она в маркдауне
            }
            return line && !line.startsWith('```') && !line.includes('```');
        });
    
        console.log('\nРезультат форматирования:');
        console.log(lines.join('\n'));
        console.log(`Количество строк в ответе: ${lines.length}`);

    return {
        description: lines[0] || 'NO_PART_NUMBER',
        source: lines[1] || 'NO_SOURCE',
            secondary_source: lines[2] || 'NO_SECOND_SOURCE'
        };
    } catch (error) {
        console.error('Ошибка при форматировании описания:', error);
        // В случае неустранимой ошибки возвращаем базовые значения
        return {
            description: 'ERROR_FORMATTING',
            source: 'NO_SOURCE',
            secondary_source: 'NO_SOURCE'
        };
    }
}

// Функция пост-обработки описания компонента
export function postProcessDescription(description: string, partNumber: string): string {
    if (!description || description === 'NO_PART_NUMBER') {
        return description;
    }

    let processedDesc = description;

    // === PART NUMBER REMOVAL ===
    if (partNumber) {
        processedDesc = processedDesc.replace(new RegExp(partNumber, 'gi'), '');
    }

    // === REMOVE UNNECESSARY WORDS ===
    processedDesc = processedDesc
        .replace(/\s*(CHIP|CHP)\s*/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    // === STANDARDIZE SPECIAL VALUES AND RANGES ===
    processedDesc = processedDesc
        // Стандартизация диапазонов напряжений и токов
        .replace(/±(\d+(?:\.\d+)?V)[-\s]*±(\d+(?:\.\d+)?V)/gi, '$1-$2')
        .replace(/±(\d+(?:\.\d+)?MA)[-\s]*±(\d+(?:\.\d+)?MA)/gi, '$1-$2')
        // Удаление ± перед значениями с единицами измерения
        .replace(/±(\d+(?:\.\d+)?(?:V|MA|MHZ|GHZ|DB|GAUSS|PPM\/K|PPM|%))/gi, '$1')
        // Обработка температурных диапазонов
        .replace(/[-\s]*(\d+)°C[-\s]*(\d+)°C/g, '$1°C$2°C')
        // Удаление оставшихся ±, +, -
        .replace(/[±+\-](?=\d)/g, '')
        .replace(/(?<=\d)[±+\-](?!\d)/g, '')
        // Удаление T&R и подобных обозначений
        .replace(/\s+(?:T&R|T\/R|TR|TAPE[- ]AND[- ]REEL)\b/gi, '');

    // === RESISTANCE STANDARDIZATION ===
    const resistanceRules: Array<[RegExp, string]> = [
        [/(\d+(?:\.\d+)?)\s*MOHM/gi, '$1M'],
        [/(\d+(?:\.\d+)?)\s*KOHM/gi, '$1K'],
        [/(\d+(?:\.\d+)?)\s*OHM/gi, '$1R'],
        [/(\d+(?:\.\d+)?)\s*Ohm/gi, '$1R'],
        [/(\d+(?:\.\d+)?)\s*ohm/gi, '$1R'],
        [/(\d+(?:\.\d+)?)\s*Ω/g, '$1R'],
        [/(\d+(?:\.\d+)?)\s*mΩ/g, '$1M'],
        [/(\d+(?:\.\d+)?)\s*kΩ/g, '$1K'],
        [/(\d+(?:\.\d+)?)\s*MΩ/g, '$1M'],
        [/mΩ/g, 'M'],
        [/Ω/g, 'R'],
        [/kΩ/g, 'K'],
        [/MΩ/g, 'M'],
        [/RDS\(ON\)\s*=\s*(\d+(?:\.\d+)?)[Ω]/gi, 'RDS$1R']
    ];

    resistanceRules.forEach(([pattern, replacement]) => {
        processedDesc = processedDesc.replace(pattern, replacement);
    });

    // === CAPACITANCE STANDARDIZATION ===
    const capacitanceRules: Array<[RegExp, string | ((match: string, value: string) => string)]> = [
        // UF/uF/μF -> MF (1:1)
        [/(\d+(?:\.\d+)?)\s*[uμ]F\b/gi, '$1MF'],
        [/(\d+(?:\.\d+)?)\s*UF\b/gi, '$1MF'],
        
        // NF/nF -> MF or PF (выбираем более короткую запись)
        [/(\d+(?:\.\d+)?)\s*[nN]F\b/gi, (match, value) => {
            const numericValue = parseFloat(value);
            // Если значение >= 100NF, конвертируем в MF (делим на 1000)
            if (numericValue >= 100) {
                return `${(numericValue / 1000).toFixed(2).replace(/\.?0+$/, '')}MF`;
            }
            // Если значение < 1NF, конвертируем в PF (умножаем на 1000)
            else if (numericValue < 1) {
                return `${(numericValue * 1000).toFixed(0)}PF`;
            }
            // Для значений от 1NF до 99NF - в PF
            else {
                return `${(numericValue * 1000).toFixed(0)}PF`;
            }
        }],
        
        // Ensure PF is uppercase
        [/(\d+(?:\.\d+)?)\s*pf\b/gi, '$1PF']
    ];

    capacitanceRules.forEach(([pattern, replacement]) => {
        if (typeof replacement === 'string') {
            processedDesc = processedDesc.replace(pattern, replacement);
        } else {
            processedDesc = processedDesc.replace(pattern, replacement);
        }
    });

    // === MOUNTING TYPE DETECTION AND ASSIGNMENT ===
    const hasExistingMountingType = /\s+(SMT|TH|THT|PTH|THROUGH[- ]HOLE)\b/i.test(processedDesc);
    
    if (!hasExistingMountingType) {
        const isThroughHole = /\b(DIP|PDIP|CDIP|CERDIP|PTH|THT|TH|THROUGH[- ]HOLE)\b/i.test(processedDesc);
        
        const smtPackages = [
            'LGA', 'BGA', 'FBGA', 'TFBGA', 'LFBGA',
            'QFN', 'TQFN', 'WQFN', 'VQFN', 'UQFN',
            'DFN', 'TDFN', 'WDFN', 'VDFN', 'UDFN',
            'QFP', 'LQFP', 'TQFP', 'VQFP', 'MQFP', 'HVQFN',
            'SON', 'WSON', 'VSON', 'USON',
            'XSON', 'DSON',
            'SOIC', 'SOP', 'SSOP', 'TSSOP', 'MSOP',
            'SOT', 'SOT23', 'SOT223', 'SOT323', 'SOT363',
            'SC70', 'SC88', 'SC70-3', 'SC70-5', 'SC70-6',
            'TO263', 'TO252', 'TO268', 'TO269',
            'PLCC'
        ].join('|');

        const isSMT = (
            new RegExp(`\\b(${smtPackages}|TDFN)[-]?\\d+\\b|\\b(${smtPackages}|TDFN)\\d+\\b`, 'i').test(processedDesc) ||
            /\b(0201|0402|0603|0805|1206|1210|1812|2010|2512)\b/i.test(processedDesc)
        );

        if (isThroughHole || (!isSMT && /\b(TO-\d+|TO\d+)\b/i.test(processedDesc))) {
            processedDesc = `${processedDesc} TH`;
        } else if (isSMT) {
            processedDesc = `${processedDesc} SMT`;
        }
        } else {
        processedDesc = processedDesc.replace(/\s+SMD\b/gi, ' SMT');
    }

    // === IC PACKAGE FORMAT RULES ===
    const packageRules: Array<[RegExp, string]> = [
        [/(\d+)-PIN\s+([A-Z]+)/gi, '$2-$1'],
        [/(\d+)-([A-Z]+)-\1/gi, '$2-$1'],
        [/(\d+)-([A-Z]+)/gi, '$2-$1'],
        [/([A-Z]+)-(\d+)-\2/gi, '$1-$2'],
        [/\b(\d+)-([A-Z]+(?:-\d+)?)\b/gi, '$2'],
        [/SMD|DIRECTFET\s+SMD/gi, 'SMT']
    ];

    packageRules.forEach(([pattern, replacement]) => {
        processedDesc = processedDesc.replace(pattern, replacement);
    });

    processedDesc = processedDesc
        .replace(/(\d+)\s+(UF|MF|PF|R|K|M|NH|MH|GHZ|MHZ|KHZ|W|MA|V|DB)/gi, '$1$2')
        .replace(/\s+/g, ' ');

    return processedDesc.trim();
}

// Обновляем функцию processRow для использования пост-обработки
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
        const formattedResult = await formatComponentDescription(searchResults, partNumber, description);
        
        // 3. Пост-обработка описания
        console.log('\nШаг 3: Пост-обработка описания');
        const processedDescription = postProcessDescription(formattedResult.description, partNumber);
        
        const result = {
            ...formattedResult,
            description: processedDescription
        };
        
        console.log('\nИтоговый результат:');
        console.log(result);
        console.log('=== Завершение обработки строки ===\n');
        
        return result;

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

    const { headers, headerRow } = await readHeaders(worksheet);
    
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
    let enrichedWorkbook: ExcelJS.Workbook | null = null;
    let enrichedSheet: ExcelJS.Worksheet | null = null;
    
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
        enrichedWorkbook = new ExcelJS.Workbook();
        enrichedSheet = enrichedWorkbook.addWorksheet('Enriched');

        if (!enrichedSheet) {
            throw new Error('Не удалось создать лист для обогащенных данных');
        }

        // Инициализируем колонки и заголовки
        const initializeColumns = (sheet: ExcelJS.Worksheet, sourceWorksheet: ExcelJS.Worksheet, headerRow: number) => {
            // Копируем заголовки и форматирование из исходного листа
            for (let colNumber = 1; colNumber <= sourceWorksheet.columnCount; colNumber++) {
                const sourceCol = sourceWorksheet.getColumn(colNumber);
                const targetCol = sheet.getColumn(colNumber);
                const headerCell = sourceWorksheet.getRow(headerRow).getCell(colNumber);
            targetCol.header = headerCell.value?.toString() || '';
            targetCol.width = sourceCol.width;
            targetCol.font = { size: 9 };
        }

            // Добавляем новые колонки
            const lastColumn = sourceWorksheet.columnCount;
        const newColumns = [
            { header: 'Enriched Description', width: 30 },
            { header: 'Primary Source', width: 30 },
            { header: 'Secondary Source', width: 30 },
            { header: 'Full Search Result', width: 40 }
        ];

            // Добавляем новые колонки после существующих
        newColumns.forEach((colConfig, index) => {
                const newColNumber = lastColumn + index + 1;
                const newCol = sheet.getColumn(newColNumber);
            newCol.header = colConfig.header;
            newCol.width = colConfig.width;
            newCol.font = { size: 9 };
        });

            // Форматируем заголовки
            const headerRowNew = sheet.getRow(1);
        headerRowNew.font = { size: 9 };

            return {
                sheetId: sheet.id,
                lastColumn: lastColumn
            };
        };

        // Инициализируем колонки и получаем ID листа и последнюю колонку
        const { sheetId: enrichedSheetId, lastColumn } = initializeColumns(enrichedSheet, worksheet, headerRow);

        const totalRows = worksheet.rowCount - headerRow;
        const firstDataRow = headerRow + 1;
        console.log(`Всего строк для обработки: ${totalRows}`);
        console.log('Первая строка с данными:', firstDataRow);

            for (let rowNumber = firstDataRow; rowNumber <= worksheet.rowCount; rowNumber++) {
            try {
                const sourceRow = worksheet.getRow(rowNumber);
                const newRow = enrichedSheet.getRow(rowNumber);
                
                // Копируем существующие значения
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
                const result = await formatComponentDescription(searchResult, partNumber, description);
                
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
                
                // Сохраняем промежуточные результаты каждые 10 строк
                if (rowNumber % 10 === 0 && enrichedWorkbook) {
                    const tempBuffer = await enrichedWorkbook.xlsx.writeBuffer();
                    if (fileId) {
                        storageManager.saveIntermediateResult(fileId, new Uint8Array(tempBuffer));
                    }
                }

            } catch (error) {
                console.error(`Ошибка при обработке строки ${rowNumber}:`, error);
                // Продолжаем обработку следующей строки
                continue;
            }
        }

        // Финальное сохранение результатов
        if (enrichedWorkbook) {
            console.log('Сохраняем финальные результаты...');
        const arrayBuffer = await enrichedWorkbook.xlsx.writeBuffer();
        console.log('Обработка файла завершена успешно');
        return new Uint8Array(arrayBuffer);
        }
        
        throw new Error('Не удалось создать обогащенный файл');
        
    } catch (error) {
        console.error('Произошла ошибка при обработке файла:', error);
        
        // Пытаемся сохранить промежуточные результаты даже при критической ошибке
        if (enrichedWorkbook) {
            try {
                console.log('Сохраняем промежуточные результаты после ошибки...');
                const arrayBuffer = await enrichedWorkbook.xlsx.writeBuffer();
                return new Uint8Array(arrayBuffer);
            } catch (saveError) {
                console.error('Ошибка при сохранении промежуточных результатов:', saveError);
                throw new Error('Не удалось сохранить промежуточные результаты');
            }
        }
        
        throw error;
    } finally {
        // Очистка ресурсов
        try {
            if (enrichedWorkbook && enrichedSheet) {
                enrichedWorkbook.removeWorksheet(enrichedSheet.id);
            }
        } catch (error) {
            console.error('Ошибка при очистке ресурсов обогащенного файла:', error);
        }

        try {
            workbook.removeWorksheet(1);
        } catch (error) {
            console.error('Ошибка при очистке основного рабочего листа:', error);
        }
    }
}

// Экспортируем функции для API
export function getAvailableFiles(fileId: string) {
    return storageManager.getAvailableFiles(fileId);
}

export function getIntermediateFile(fileId: string, fileName: string) {
    return storageManager.getFile(fileId, fileName);
}

