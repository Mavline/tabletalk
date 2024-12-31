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
    private readonly uploadsDir = 'uploads';
    private readonly processedDir = 'processed';
    private readonly storageDir = './table_storage';
    private readonly maxFiles = 10;
    private readonly cleanupInterval = 24 * 60 * 60 * 1000; // 24 часа

    private constructor() {
        if (!fs.existsSync(this.storageDir)) {
            fs.mkdirSync(this.storageDir, { recursive: true });
        }
        this.cleanupOldFiles();
        setInterval(() => this.cleanup(), this.cleanupInterval);
    }

    public static getInstance(): StorageManager {
        if (!StorageManager.instance) {
            StorageManager.instance = new StorageManager();
        }
        return StorageManager.instance;
    }

    private ensureDirectories() {
        [this.uploadsDir, this.processedDir, this.storageDir].forEach(dir => {
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
            }
        });
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
        const metadataPath = path.join(this.storageDir, `${fileId}.json`);
        await fs.promises.writeFile(metadataPath, JSON.stringify(metadata, null, 2));

        return fileId;
    }

    public async saveProcessedFile(fileId: string, buffer: Buffer): Promise<string> {
        const metadataPath = path.join(this.storageDir, `${fileId}.json`);
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
        const metadataPath = path.join(this.storageDir, `${fileId}.json`);
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
        const metadataPath = path.join(this.storageDir, `${fileId}.json`);
        const metadata: FileMetadata = JSON.parse(await fs.promises.readFile(metadataPath, 'utf-8'));
        return metadata.chatHistory;
    }

    public async getProcessedFilePath(fileId: string): Promise<string> {
        const metadataPath = path.join(this.storageDir, `${fileId}.json`);
        const metadata: FileMetadata = JSON.parse(await fs.promises.readFile(metadataPath, 'utf-8'));
        return path.join(this.processedDir, metadata.processedName);
    }

    private async cleanup() {
        try {
            const files = await fs.promises.readdir(this.storageDir);
            const metadataFiles = files.filter(f => f.endsWith('.json'));

            if (metadataFiles.length <= this.maxFiles) return;

            const metadataWithStats = await Promise.all(
                metadataFiles.map(async (file) => {
                    const filePath = path.join(this.storageDir, file);
                    const metadata: FileMetadata = JSON.parse(
                        await fs.promises.readFile(filePath, 'utf-8')
                    );
                    return { file, metadata };
                })
            );

            // Сортируем по времени последнего доступа
            metadataWithStats.sort(
                (a, b) => a.metadata.lastAccess.getTime() - b.metadata.lastAccess.getTime()
            );

            // Удаляем старые файлы
            const filesToRemove = metadataWithStats.slice(0, metadataWithStats.length - this.maxFiles);
            
            for (const { file, metadata } of filesToRemove) {
                const fileId = file.replace('.json', '');
                const uploadPath = path.join(this.uploadsDir, fileId);
                const processedPath = path.join(this.processedDir, metadata.processedName);
                const metadataPath = path.join(this.storageDir, file);

                await Promise.all([
                    fs.promises.unlink(uploadPath).catch(() => {}),
                    fs.promises.unlink(processedPath).catch(() => {}),
                    fs.promises.unlink(metadataPath).catch(() => {})
                ]);
            }
        } catch (error) {
            console.error('Ошибка при очистке старых файлов:', error);
        }
    }

    private cleanupOldFiles() {
        try {
            const files = fs.readdirSync(this.storageDir);
            if (files.length > this.maxFiles) {
                // Сортируем файлы по времени создания
                const sortedFiles = files
                    .map(file => ({
                        name: file,
                        time: fs.statSync(path.join(this.storageDir, file)).birthtime.getTime()
                    }))
                    .sort((a, b) => a.time - b.time);

                // Удаляем самые старые файлы
                const filesToDelete = sortedFiles.slice(0, files.length - this.maxFiles);
                filesToDelete.forEach(file => {
                    fs.unlinkSync(path.join(this.storageDir, file.name));
                    console.log(`Удален старый файл: ${file.name}`);
                });
            }
        } catch (error) {
            console.error('Ошибка при очистке старых файлов:', error);
        }
    }
} 