document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('bomFile');
    const statusDiv = document.getElementById('processingStatus');
    const progressBar = document.querySelector('.progress');
    const previewTable = document.getElementById('previewTable');
    const downloadBtn = document.getElementById('downloadBtn');

    let processedData = null;

    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('bomFile', file);

        try {
            statusDiv.textContent = 'Загрузка файла...';
            progressBar.style.width = '10%';
            previewTable.style.display = 'none';
            downloadBtn.style.display = 'none';

            // Отправляем файл на сервер
            const response = await fetch('/api/process-bom', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error('Ошибка загрузки файла');
            }

            // Получаем поток событий обработки
            const reader = response.body.getReader();
            
            while (true) {
                const {value, done} = await reader.read();
                if (done) break;
                
                const text = new TextDecoder().decode(value);
                const data = JSON.parse(text);

                if (data.type === 'progress') {
                    // Обновляем прогресс
                    statusDiv.textContent = `Обработано ${data.current} из ${data.total} строк...`;
                    const progress = (data.current / data.total) * 100;
                    progressBar.style.width = `${progress}%`;
                } else if (data.type === 'preview') {
                    // Показываем предпросмотр изменений
                    const tbody = previewTable.querySelector('tbody');
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${data.before}</td>
                        <td>${data.after}</td>
                        <td><a href="${data.source}" target="_blank">Источник</a></td>
                    `;
                    tbody.appendChild(row);
                    previewTable.style.display = 'table';
                }
            }

            // Завершение обработки
            statusDiv.textContent = 'Обработка завершена!';
            progressBar.style.width = '100%';
            downloadBtn.style.display = 'block';
            
        } catch (error) {
            statusDiv.textContent = `Ошибка: ${error.message}`;
            progressBar.style.width = '0%';
        }
    });

    downloadBtn.addEventListener('click', async () => {
        try {
            const response = await fetch('/api/download-result');
            const blob = await response.blob();
            
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const fileNameWithoutExt = fileInput.files[0].name.replace('.xlsx', '');
            a.download = `${fileNameWithoutExt}_enreached.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } catch (error) {
            statusDiv.textContent = `Ошибка скачивания: ${error.message}`;
        }
    });
});