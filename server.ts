import express, { Request, Response } from 'express';
import multer from 'multer';
import cors from 'cors';
import { processExcelBuffer, askLLM } from './bomEnricher';
import { Server as HttpServer } from 'http';
import { WebSocket, WebSocketServer } from 'ws';
import { StorageManager } from './storage';
import * as fs from 'fs';
import ExcelJS from 'exceljs';

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
        res.status(400).json({ error: 'Файл не загружен' });
        return;
    }

    if (!req.file.originalname.match(/\.(xlsx|xls)$/i)) {
        res.status(400).json({ error: 'Поддерживаются только файлы Excel (.xlsx, .xls)' });
        return;
    }

    try {
        console.log(`Начинаем обработку файла: ${req.file.originalname}`);
        
        // Сохраняем загруженный файл
        const fileId = await storageManager.saveUploadedFile(req.file.buffer, req.file.originalname);

        // Сохраняем начальное сообщение LLM в историю чата
        await storageManager.addChatMessage(fileId, 'assistant', 
            `Я анализирую таблицу "Bill of Materials" и помогаю дополнить информацию о компонентах и их названиях посредством поиска в источниках на сайтах:
- https://www.digikey.co.il/en
- https://www.mouser.com/
- https://www.datasheets360.com/

Я могу:
1. Определить недостающую информацию в описаниях
2. Найти компоненты по парт-номеру
3. Дополнить описания, сохраняя существующий стиль
4. Ответить на вопросы о компонентах

Чем могу помочь?`
        );

        // Получаем количество строк в файле
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(1);
        const totalRows = worksheet ? worksheet.rowCount - 1 : 0;

        const processedBuffer = await processExcelBuffer(
            req.file.buffer,
            (current: number, total: number) => {
                console.log(`Прогресс: ${current}/${total} строк`);
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
                console.log('Предпросмотр изменений:');
                console.log('До:', before);
                console.log('После:', after);
                console.log('Источник:', source);
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
            `Файл успешно обработан. Обработано строк: ${totalRows}. 
Вы можете задавать вопросы о компонентах или запросить дополнительную информацию.`
        );

        console.log('Отправляем обработанный файл клиенту');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${processedName}`);
        res.send(Buffer.from(processedBuffer));

    } catch (error) {
        console.error('Ошибка при обработке файла:', error);
        const errorMessage = error instanceof Error ? error.message : 'Неизвестная ошибка';
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
            res.status(400).json({ error: 'Вопрос не указан' });
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
            error: error instanceof Error ? error.message : 'Ошибка при обращении к LLM'
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
            error: error instanceof Error ? error.message : 'История чата не найдена'
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
            error: error instanceof Error ? error.message : 'Файл не найден'
        });
    }
});

// Обработка WebSocket подключений
wss.on('connection', (ws: WebSocket) => {
    console.log('Новое WebSocket подключение');
    
    ws.on('close', () => {
        console.log('WebSocket подключение закрыто');
    });
});

const PORT = process.env.PORT || 3002;
server.listen(PORT, () => {
    console.log(`Сервер запущен на порту ${PORT}`);
}); 