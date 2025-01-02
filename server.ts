import express, { Request, Response } from 'express';
import multer from 'multer';
import cors from 'cors';
import { processExcelBuffer, askLLM, cleanupTempFiles } from './bomEnricher';
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

    try {
        console.log(`Starting file processing: ${req.file.originalname}`);
        
        // Сохраняем загруженный файл
        const fileId = await storageManager.saveUploadedFile(req.file.buffer, req.file.originalname);

        // Сохраняем начальное сообщение LLM в историю чата
        await storageManager.addChatMessage(fileId, 'assistant', 
            `I am analyzing the "Bill of Materials" table and helping to supplement information about components and their names by searching in sources on websites:

I can:
1. Identify missing information in descriptions
2. Find components by part number
3. Supplement descriptions while preserving the existing style
4. Answer questions about components

How can I help?`
        );

        // Получаем количество строк в файле
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(1);
        const totalRows = worksheet ? worksheet.rowCount - 1 : 0;

        const processedBuffer = await processExcelBuffer(
            req.file.buffer,
            (current: number, total: number) => {
                console.log(`Progress: ${current}/${total} rows`);
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
                console.log('Preview of changes:');
                console.log('Before:', before);
                console.log('After:', after);
                console.log('Source:', source);
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

        // Сохраняем обработанный файл
        const processedName = await storageManager.saveProcessedFile(fileId, Buffer.from(processedBuffer));

        // Добавляем сообщение о завершении обработки
        await storageManager.addChatMessage(fileId, 'assistant', 
            `File processed successfully. Processed rows: ${totalRows}. 
You can ask questions about components or request additional information.`
        );

        console.log('Sending processed file to client');
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

// Маршрут для диалога с LLM
app.post('/api/ask-llm', async (req: Request, res: Response): Promise<void> => {
    try {
        const { question, fileId } = req.body;
        if (!question) {
            res.status(400).json({ error: 'Question not specified' });
            return;
        }

        // Сохраняем вопрос пользователя
        if (fileId) {
            await storageManager.addChatMessage(fileId, 'user', question);
        }

        const answer = await askLLM(question, fileId);

        // Сохраняем ответ LLM
        if (fileId) {
            await storageManager.addChatMessage(fileId, 'assistant', answer);
        }

        res.json({ answer });
    } catch (error) {
        res.status(500).json({ 
            error: error instanceof Error ? error.message : 'Error contacting LLM'
        });
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
    console.log('New WebSocket connection');
    
    // Очищаем временные файлы при новом подключении
    console.log('Cleaning up temporary files...');
    cleanupTempFiles();
    console.log('Cleanup completed');

    ws.on('message', async (message: string) => {
        try {
            const data = JSON.parse(message);
            // ... rest of the WebSocket message handling ...
        } catch (error) {
            console.error('Error processing WebSocket message:', error);
            ws.send(JSON.stringify({ error: 'Error processing message' }));
        }
    });

    ws.on('close', () => {
        console.log('WebSocket connection closed');
    });
});

const PORT = process.env.PORT || 3002;
server.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
}); 
