import express, { Request, Response } from 'express';
import multer from 'multer';
import cors from 'cors';
import { processExcelBuffer, getFileHeaders, cleanupTempFiles, getSheetNames, getFileHeadersFromSheet } from './bomEnricher';
import { Server as HttpServer } from 'http';
import { WebSocket, WebSocketServer } from 'ws';
import { StorageManager } from './storage';
import * as fs from 'fs';
import ExcelJS from 'exceljs';
import dotenv from 'dotenv';

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
app.post('/api/get-headers', async (req: Request, res: Response): Promise<void> => {
    try {
        const { fileId, sheet } = req.body;
        
        if (!fileId) {
            res.status(400).json({ error: 'Не указан идентификатор файла' });
            return;
        }

        // Получаем файл из хранилища
        const filePath = await storageManager.getUploadedFilePath(fileId);
        const buffer = await fs.promises.readFile(filePath);

        let headers;
        if (sheet) {
            headers = await getFileHeadersFromSheet(buffer, sheet);
        } else {
            headers = await getFileHeaders(buffer);
        }

        res.json({ headers });
    } catch (error) {
        console.error('Ошибка при чтении заголовков:', error);
        res.status(500).json({ error: 'Ошибка при чтении заголовков файла' });
    }
});

// Добавляем новый маршрут для получения списка листов
app.post('/api/get-sheets', upload.single('file'), async (req: Request, res: Response): Promise<void> => {
    console.log('=== Получен запрос на получение списка листов ===');
    try {
        if (!req.file) {
            console.error('Файл не был загружен');
            res.status(400).json({ error: 'Файл не был загружен' });
            return;
        }

        // Сохраняем файл во временное хранилище
        const fileId = await storageManager.saveUploadedFile(req.file.buffer, req.file.originalname);
        
        // Получаем список листов
        const sheets = await getSheetNames(req.file.buffer);

        if (!sheets || sheets.length === 0) {
            console.error('Листы не найдены в файле');
            res.status(400).json({ error: 'Листы не найдены в файле' });
            return;
        }

        res.json({ sheets, fileId });
        
        // Очищаем память
        req.file.buffer = Buffer.from([]);
    } catch (error) {
        console.error('Ошибка при получении списка листов:', error);
        res.status(500).json({ error: 'Ошибка при получении списка листов' });
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

    if (isNaN(partNumberColumn) || isNaN(descriptionColumn)) {
        res.status(400).json({ error: 'Invalid column indices' });
        return;
    }

    try {
        // Сохраняем файл во временное хранилище
        const fileId = await storageManager.saveUploadedFile(req.file.buffer, req.file.originalname);

        const processedBuffer = await processExcelBuffer(
            req.file.buffer,
            sheetName,
            partNumberColumn,
            descriptionColumn,
            (current: number, total: number) => {
                wss.clients.forEach((client: WebSocket) => {
                    if (client.readyState === WebSocket.OPEN) {
                        client.send(JSON.stringify({
                            type: 'progress',
                            fileId,
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
                            fileId,
                            before,
                            after,
                            source
                        }));
                    }
                });
            },
            fileId
        );

        // Сохраняем только обработанный лист
        const processedName = await storageManager.saveProcessedFile(fileId, Buffer.from(processedBuffer));

        console.log('Отправляем обработанный файл клиенту');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${processedName}`);
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

// Маршрут для получения истории чата
app.get('/api/chat-history/:fileId', async (req: Request, res: Response): Promise<void> => {
    try {
        const { fileId } = req.params;
        const history = await storageManager.getChatHistory(fileId);
        res.json({ history });
    } catch (error) {
        res.status(404).json({ 
            error: error instanceof Error ? error.message : 'Chat history not found'
        });
    }
});

// Маршрут для получения обработанного файла
app.get('/api/file/:fileId', async (req: Request, res: Response): Promise<void> => {
    try {
        const { fileId } = req.params;
        const filePath = await storageManager.getProcessedFilePath(fileId);
        const fileContent = await fs.promises.readFile(filePath);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(fileContent);
    } catch (error) {
        res.status(404).json({ 
            error: error instanceof Error ? error.message : 'File not found'
        });
    }
});

// Обработка WebSocket подключений
wss.on('connection', (ws: WebSocket) => {
    console.log('WebSocket подключение установлено');
    cleanupTempFiles();

    ws.on('message', async (message: string) => {
        try {
            const data = JSON.parse(message);
            
            if (data.type === 'progress') {
                // Обработка прогресса
                console.log('Progress:', data.current, '/', data.total);
            }
            // Удаляем обработку сообщений чата
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
