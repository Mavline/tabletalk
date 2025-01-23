import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';

export class StorageManager {
    private static instance: StorageManager;
    private readonly maxFiles = 10; // Максимум 10 файлов
    private readonly resultsDir = path.join(process.cwd(), 'storage', 'results');
    private readonly tempDir = path.join(process.cwd(), 'storage', 'temp');
    private fileCache: Map<string, { buffer: Buffer, timestamp: number }> = new Map();
    private readonly cacheTimeout = 5 * 60 * 1000; // 5 минут
    private readonly MAX_FILES = 10;

    private constructor() {
        // Создаем директории при инициализации
        [this.resultsDir, this.tempDir].forEach(dir => {
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
            }
        });

        // Запускаем периодическую очистку кэша
        setInterval(() => this.cleanupCache(), this.cacheTimeout);
    }

    static getInstance(): StorageManager {
        if (!StorageManager.instance) {
            StorageManager.instance = new StorageManager();
        }
        return StorageManager.instance;
    }

    // Сохранение выбранного листа во временный файл
    async cacheSheet(buffer: Buffer, sheetName: string): Promise<string> {
        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            
            const sheet = workbook.getWorksheet(sheetName);
            if (!sheet) {
                throw new Error('Sheet not found');
            }

            // Создаем новый файл только с нужным листом
            const newWorkbook = new ExcelJS.Workbook();
            const newSheet = newWorkbook.addWorksheet(sheetName);
            
            // Копируем данные
            sheet.eachRow((row, rowIndex) => {
                const newRow = newSheet.getRow(rowIndex);
                row.eachCell((cell, colIndex) => {
                    newRow.getCell(colIndex).value = cell.value;
                });
                newRow.commit();
            });

            // Генерируем ID для файла
            const fileId = `${Date.now()}_${sheetName}`;
            const tempBuffer = await newWorkbook.xlsx.writeBuffer();
            
            // Сохраняем в кэш
            this.fileCache.set(fileId, {
                buffer: Buffer.from(tempBuffer),
                timestamp: Date.now()
            });

            return fileId;
        } catch (error) {
            console.error('Ошибка при кэшировании листа:', error);
            throw error;
        }
    }

    // Получение кэшированного листа
    getCachedSheet(fileId: string): Buffer | null {
        const cached = this.fileCache.get(fileId);
        if (!cached) return null;

        // Обновляем timestamp при обращении
        cached.timestamp = Date.now();
        return cached.buffer;
    }

    // Очистка старых файлов из кэша
    private cleanupCache(): void {
        const now = Date.now();
        for (const [fileId, { timestamp }] of this.fileCache.entries()) {
            if (now - timestamp > this.cacheTimeout) {
                this.fileCache.delete(fileId);
            }
        }
    }

    // Сохранение готового файла
    async saveProcessedFile(originalName: string, data: Uint8Array): Promise<string> {
        try {
            const baseName = path.parse(originalName).name;
            
            // Формируем базовое имя файла
            let baseFileName = `${baseName}_enreached.xlsx`;
            let fileName = baseFileName;
            let counter = 1;

            // Проверяем существование файла и добавляем номер, если нужно
            while (fs.existsSync(path.join(this.resultsDir, fileName))) {
                fileName = `${baseName}_enreached(${counter}).xlsx`;
                counter++;
            }
            
            const filePath = path.join(this.resultsDir, fileName);
            fs.writeFileSync(filePath, Buffer.from(data));

            // Проверяем и очищаем старые файлы
            this.cleanupOldFiles();
            
            return fileName;
        } catch (error) {
            console.error('Error saving file:', error);
            throw error;
        }
    }

    private cleanupOldFiles(): void {
        const files = fs.readdirSync(this.resultsDir);
        if (files.length > this.MAX_FILES) {
            const sortedFiles = files
                .map(file => ({
                    name: file,
                    path: path.join(this.resultsDir, file),
                    stats: fs.statSync(path.join(this.resultsDir, file))
                }))
                .sort((a, b) => b.stats.mtimeMs - a.stats.mtimeMs);

            // Keep only MAX_FILES most recent files
            sortedFiles.slice(this.MAX_FILES).forEach(file => {
                try {
                    fs.unlinkSync(file.path);
                } catch (error) {
                    console.error('Error deleting file:', error);
                }
            });
        }
    }
} 