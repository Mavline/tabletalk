// Backend code for executing JavaScript and Python
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const { processExcelFile } = require('./bomEnricher');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.static(__dirname));

// Маршрут для загрузки и обработки файла
app.post('/api/process-bom', upload.single('bomFile'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'Файл не загружен' });
    }

    // Настраиваем SSE для отправки прогресса
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    try {
        const inputPath = req.file.path;
        const outputPath = path.join(__dirname, 'processed', `${Date.now()}_enriched.xlsx`);

        // Функция для отправки прогресса клиенту
        const sendProgress = (current, total) => {
            res.write(JSON.stringify({
                type: 'progress',
                current,
                total
            }));
        };

        // Функция для отправки предпросмотра изменений
        const sendPreview = (before, after, source) => {
            res.write(JSON.stringify({
                type: 'preview',
                before,
                after,
                source
            }));
        };

        // Обработка файла
        await processExcelFile(inputPath, outputPath, sendProgress, sendPreview);

        res.end();
    } catch (error) {
        res.write(JSON.stringify({ type: 'error', message: error.message }));
        res.end();
    }
});

// Маршрут для скачивания обработанного файла
app.get('/api/download-result', (req, res) => {
    // Здесь нужно реализовать отправку последнего обработанного файла
    const lastProcessedFile = getLastProcessedFile(); // Реализовать эту функцию
    if (lastProcessedFile) {
        res.download(lastProcessedFile);
    } else {
        res.status(404).json({ error: 'Файл не найден' });
    }
});

const PORT = process.env.PORT || 3002;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});