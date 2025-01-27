import { Command } from 'commander';
import { processExcelBuffer } from './bomEnricher';
import * as path from 'path';
import * as fs from 'fs';

interface EnrichOptions {
    output: string;
}

const program = new Command();

program
    .name('bom-enricher')
    .description('Инструмент для обогащения описаний компонентов в BOM файлах')
    .version('1.0.0');

program
    .command('enrich')
    .description('Обогатить BOM файл')
    .argument('<input>', 'Путь к входному Excel файлу')
    .option('-o, --output <path>', 'Путь к выходному файлу', 'enriched_bom.xlsx')
    .action(async (input: string, options: EnrichOptions) => {
        console.log('Начинаем обработку файла...');
        
        try {
            const inputPath = path.resolve(input);
            const outputPath = path.resolve(options.output);

            const buffer = await fs.promises.readFile(inputPath);
            const processedBuffer = await processExcelBuffer(
                buffer,
                'Sheet1',
                1,
                2,
                (current: number, total: number) => {
                    const percent = Math.round((current / total) * 100);
                    process.stdout.write(`\rПрогресс: ${percent}% (${current}/${total})`);
                },
                (before: string, after: string, source: string) => {
                    console.log('\nОбновлено:');
                    console.log(`Было:  ${before}`);
                    console.log(`Стало: ${after}`);
                    console.log(`Источник: ${source}\n`);
                }
            );

            await fs.promises.writeFile(outputPath, Buffer.from(processedBuffer));
            console.log('\nГотово! Результат сохранен в:', outputPath);
        } catch (error: unknown) {
            if (error instanceof Error) {
                console.error('\nОшибка:', error.message);
            } else {
                console.error('\nНеизвестная ошибка:', error);
            }
            process.exit(1);
        }
    });

program.parse(); 