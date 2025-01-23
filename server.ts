import express, { Request, Response } from 'express';
import multer from 'multer';
import cors from 'cors';
import { processExcelBuffer, getFileHeaders, getSheetNames, getFileHeadersFromSheet } from './bomEnricher';
import { Server as HttpServer } from 'http';
import { WebSocket, WebSocketServer } from 'ws';
import { StorageManager } from './storage';
import * as fs from 'fs';
import ExcelJS from 'exceljs';
import dotenv from 'dotenv';
import path from 'path';

dotenv.config();

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const server = new HttpServer(app);
const wss = new WebSocketServer({ server });
const storageManager = StorageManager.getInstance();

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Маршрут для получения заголовков файла
app.post('/api/get-headers', upload.single('file'), async (req: Request, res: Response): Promise<void> => {
    try {
        const { sheet } = req.body;
        const buffer = req.file?.buffer;

        if (!buffer) {
            res.status(400).json({ error: 'Файл не предоставлен' });
            return;
        }

        // Кэшируем выбранный лист
        const fileId = await storageManager.cacheSheet(buffer, sheet);

        let headers;
        const cachedBuffer = storageManager.getCachedSheet(fileId);
        if (cachedBuffer) {
            headers = await getFileHeadersFromSheet(cachedBuffer, sheet);
        } else {
            headers = await getFileHeadersFromSheet(buffer, sheet);
        }

        res.json({ headers, fileId });
    } catch (error) {
        console.error('Ошибка при чтении заголовков:', error);
        res.status(500).json({ error: 'Ошибка при чтении заголовков файла' });
    }
});

// Маршрут для получения списка листов
app.post('/api/get-sheets', upload.single('file'), async (req: Request, res: Response): Promise<void> => {
    console.log('=== Получен запрос на получение списка листов ===');
    try {
        if (!req.file) {
            console.error('Файл не был загружен');
            res.status(400).json({ error: 'Файл не был загружен' });
            return;
        }

        // Получаем список листов
        const sheets = await getSheetNames(req.file.buffer);

        if (!sheets || sheets.length === 0) {
            console.error('Листы не найдены в файле');
            res.status(400).json({ error: 'Листы не найдены в файле' });
            return;
        }

        res.json({ sheets });
        
        // Очищаем память
        req.file.buffer = Buffer.from([]);
    } catch (error) {
        console.error('Ошибка при получении списка листов:', error);
        res.status(500).json({ error: 'Ошибка при получении списка листов' });
    }
});

// Маршрут для получения списка готовых файлов
app.get('/api/processed-files', async (req: Request, res: Response): Promise<void> => {
    try {
        const resultsDir = path.join(process.cwd(), 'storage', 'results');
        if (!fs.existsSync(resultsDir)) {
            res.json({ files: [] });
            return;
        }

        const files = fs.readdirSync(resultsDir)
            .filter(file => file.endsWith('.xlsx'))
            .map(file => ({
                name: file,
                size: fs.statSync(path.join(resultsDir, file)).size
            }))
            .sort((a, b) => b.size - a.size)
            .slice(0, 20); // Показываем только последние 20 файлов

        res.json({ files });
    } catch (error) {
        console.error('Error getting file list:', error);
        res.status(500).json({ error: 'Failed to get file list' });
    }
});

// Маршрут для скачивания готового файла
app.get('/api/download-processed/:filename', async (req: Request, res: Response): Promise<void> => {
    try {
        const { filename } = req.params;
        const filePath = path.join(process.cwd(), 'storage', 'results', filename);
        
        if (!fs.existsSync(filePath)) {
            res.status(404).json({ error: 'Файл не найден' });
            return;
        }

        res.download(filePath);
    } catch (error) {
        console.error('Ошибка при скачивании файла:', error);
        res.status(500).json({ error: 'Ошибка при скачивании файла' });
    }
});

// Маршрут для обработки файла
app.post('/process', upload.single('file'), async (req: Request, res: Response): Promise<void> => {
    if (!req.file?.buffer) {
        res.status(400).json({ error: 'File not uploaded' });
        return;
    }

    if (!req.file.originalname.match(/\.(xlsx|xls)$/i)) {
        res.status(400).json({ error: 'Only Excel files (.xlsx, .xls) are supported' });
        return;
    }

    const partNumberColumn = parseInt(req.body.partNumberColumn);
    const descriptionColumn = parseInt(req.body.descriptionColumn);
    const sheetName = req.body.sheet;
    const fileId = req.body.fileId;

    if (isNaN(partNumberColumn) || isNaN(descriptionColumn)) {
        res.status(400).json({ error: 'Invalid column indices' });
        return;
    }

    try {
        // Пытаемся использовать кэшированный лист
        let buffer = fileId ? storageManager.getCachedSheet(fileId) : null;
        if (!buffer) {
            buffer = req.file.buffer;
        }

        const processedBuffer = await processExcelBuffer(
            buffer,
            sheetName,
            partNumberColumn,
            descriptionColumn,
            (current: number, total: number) => {
                wss.clients.forEach((client: WebSocket) => {
                    if (client.readyState === WebSocket.OPEN) {
                        client.send(JSON.stringify({
                            type: 'progress',
                            current,
                            total
                        }));
                    }
                });
            },
            (before: string, after: string, source: string) => {
                wss.clients.forEach((client: WebSocket) => {
                    if (client.readyState === WebSocket.OPEN) {
                        client.send(JSON.stringify({
                            type: 'preview',
                            before,
                            after,
                            source
                        }));
                    }
                });
            }
        );

        // Сохраняем готовый файл
        const processedName = await storageManager.saveProcessedFile(req.file.originalname, processedBuffer);

        console.log('Отправляем обработанный файл клиенту');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${processedName}"`);
        res.send(Buffer.from(processedBuffer));

    } catch (error) {
        console.error('Error processing file:', error);
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        wss.clients.forEach((client: WebSocket) => {
            if (client.readyState === WebSocket.OPEN) {
                client.send(JSON.stringify({
                    type: 'error',
                    message: errorMessage
                }));
            }
        });
        res.status(500).json({ error: errorMessage });
    }
});

// WebSocket подключение
wss.on('connection', (ws: WebSocket) => {
    console.log('WebSocket подключение установлено');

    ws.on('message', async (message: string) => {
        try {
            const data = JSON.parse(message);
            
            if (data.type === 'progress') {
                console.log('Progress:', data.current, '/', data.total);
            }
        } catch (error) {
            console.error('Ошибка при обработке WebSocket сообщения:', error);
            ws.send(JSON.stringify({ error: 'Ошибка при обработке сообщения' }));
        }
    });

    ws.on('close', () => {
        console.log('WebSocket подключение закрыто');
    });
});

const PORT = process.env.PORT || 3002;
server.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
}); 
