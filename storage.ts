import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';

export interface ChatMessage {
    role: 'user' | 'assistant';
    content: string;
    timestamp: Date;
}

interface FileMetadata {
    originalName: string;
    processedName: string;
    uploadTime: Date;
    lastAccess: Date;
    chatHistory: ChatMessage[];
}

export class StorageManager {
    private static instance: StorageManager;
    private uploadsDir: string;
    private processedDir: string;
    private tableStorageDir: string;

    private constructor() {
        this.uploadsDir = path.join(process.cwd(), 'uploads');
        this.processedDir = path.join(process.cwd(), 'processed');
        this.tableStorageDir = path.join(process.cwd(), 'table_storage');

        // Создаем директории если они не существуют
        [this.uploadsDir, this.processedDir, this.tableStorageDir].forEach(dir => {
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
            }
        });
    }

    public static getInstance(): StorageManager {
        if (!StorageManager.instance) {
            StorageManager.instance = new StorageManager();
        }
        return StorageManager.instance;
    }

    private generateFileId(originalName: string): string {
        const timestamp = Date.now();
        const hash = crypto.createHash('md5')
            .update(`${originalName}_${timestamp}`)
            .digest('hex');
        return hash;
    }

    public async saveUploadedFile(buffer: Buffer, originalName: string): Promise<string> {
        const fileId = this.generateFileId(originalName);
        const metadata: FileMetadata = {
            originalName,
            processedName: '',
            uploadTime: new Date(),
            lastAccess: new Date(),
            chatHistory: []
        };

        // Сохраняем файл
        const uploadPath = path.join(this.uploadsDir, fileId);
        await fs.promises.writeFile(uploadPath, buffer);

        // Сохраняем метаданные
        const metadataPath = path.join(this.tableStorageDir, `${fileId}.json`);
        await fs.promises.writeFile(metadataPath, JSON.stringify(metadata, null, 2));

        return fileId;
    }

    public async saveProcessedFile(fileId: string, buffer: Buffer): Promise<string> {
        const metadataPath = path.join(this.tableStorageDir, `${fileId}.json`);
        const metadata: FileMetadata = JSON.parse(await fs.promises.readFile(metadataPath, 'utf-8'));

        const processedName = `${Date.now()}_enriched.xlsx`;
        const processedPath = path.join(this.processedDir, processedName);
        
        await fs.promises.writeFile(processedPath, buffer);
        
        metadata.processedName = processedName;
        metadata.lastAccess = new Date();
        
        await fs.promises.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
        
        return processedName;
    }

    public async addChatMessage(fileId: string, role: 'user' | 'assistant', content: string) {
        const metadataPath = path.join(this.tableStorageDir, `${fileId}.json`);
        const metadata: FileMetadata = JSON.parse(await fs.promises.readFile(metadataPath, 'utf-8'));

        metadata.chatHistory.push({
            role,
            content,
            timestamp: new Date()
        });
        metadata.lastAccess = new Date();

        await fs.promises.writeFile(metadataPath, JSON.stringify(metadata, null, 2));
    }

    public async getChatHistory(fileId: string): Promise<ChatMessage[]> {
        const metadataPath = path.join(this.tableStorageDir, `${fileId}.json`);
        const metadata: FileMetadata = JSON.parse(await fs.promises.readFile(metadataPath, 'utf-8'));
        return metadata.chatHistory;
    }

    public async getProcessedFilePath(fileId: string): Promise<string> {
        const metadataPath = path.join(this.tableStorageDir, `${fileId}.json`);
        const metadata: FileMetadata = JSON.parse(await fs.promises.readFile(metadataPath, 'utf-8'));
        return path.join(this.processedDir, metadata.processedName);
    }

    public async cleanup() {
        try {
            // Очищаем все директории
            const cleanDir = async (dir: string) => {
                if (!fs.existsSync(dir)) return;
                const files = await fs.promises.readdir(dir);
                await Promise.all(
                    files.map(file => 
                        fs.promises.unlink(path.join(dir, file)).catch(() => {})
                    )
                );
            };

            // Очищаем все три директории
            await Promise.all([
                cleanDir(this.uploadsDir),
                cleanDir(this.processedDir),
                cleanDir(this.tableStorageDir)
            ]);

            // Создаем директории если они не существуют
            [this.uploadsDir, this.processedDir, this.tableStorageDir].forEach(dir => {
                if (!fs.existsSync(dir)) {
                    fs.mkdirSync(dir, { recursive: true });
                }
            });
        } catch (error) {
            console.error('Ошибка при очистке файлов:', error);
        }
    }

    // Получение пути к загруженному файлу
    async getUploadedFilePath(fileId: string): Promise<string> {
        const filePath = path.join(this.uploadsDir, fileId);
        if (!fs.existsSync(filePath)) {
            throw new Error('Файл не найден');
        }
        return filePath;
    }
} 