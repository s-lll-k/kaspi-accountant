import express from 'express';
import { Telegraf, session } from 'telegraf';
import xlsx from 'xlsx';
import fs from 'fs';
import fetch from 'node-fetch';
import dotenv from 'dotenv';

dotenv.config()

const app = express();
const bot = new Telegraf(process.env.TELEGRAF_TOKEN); // Замените на токен вашего бота

/**
 * хранит себестоимость товаров
 */
let productPrices;

bot.use(session());

// Установка команд для бота
bot.telegram.setMyCommands([
    { command: 'setprices', description: 'Установить цены на продукты' },
    { command: 'countresult', description: 'Вычислить результат' },
]);

// Команда для установки цен на продукты
bot.command('setprices', ctx => {
    session.waitingForCount = false; // Устанавливаем флаг ожидания цен
    session.waitingForPrices = true; // Устанавливаем флаг ожидания таблицы kaspi продаж
    ctx.reply('Скинь таблицу c себестоимостью товаров');
});

// Команда для подсчета результата
bot.command('countresult', ctx => {
    session.waitingForPrices = false; // Устанавливаем флаг ожидания цен
    session.waitingForCount = true; // Устанавливаем флаг ожидания таблицы себестоимостей
    ctx.reply('Скинь таблицу из kaspi pay');
});

// Хендлер для обработки следующего сообщения от пользователя
bot.on('text', (ctx) => {
    ctx.reply('Тиx...Тих..Чш...Ч.Ч..ч.ч.ч. просто используй команды в меню.');
});

// Обработка ошибок
bot.catch((err, ctx) => {
    console.error(`Ошибка в обработке апдейта для ${ctx.updateType}`, err);
    ctx.reply('Произошла ошибка. Попробуйте еще раз.');
});

// Хендлер для обработки файлов .xlsx
bot.on('document', async (ctx) => {
    try {
        // Проверяем, что файл является .xlsx
        const fileExtension = ctx.message.document.file_name.split('.').pop().toLowerCase();
        if (fileExtension !== 'xlsx') {
            await ctx.reply('Пожалуйста, отправьте файл в формате .xlsx');
            return;
        }

        if (session.waitingForPrices) {
            const fileId = ctx.message.document.file_id;
            const fileLink = await ctx.telegram.getFileLink(fileId);
            const response = await fetch(fileLink.href);
            const buffer = await response.buffer();

            fs.writeFileSync(`uploads/${ctx.message.document.file_name}`, buffer);
            
            const workbook = xlsx.readFile(`uploads/${ctx.message.document.file_name}`);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = xlsx.utils.sheet_to_json(sheet, {
                header: 1,
                raw: true
            });

            if (data.length === 0 || !data[0].length) {
                await ctx.reply('Файл пуст или содержит неверный формат данных');
                fs.unlinkSync(`uploads/${ctx.message.document.file_name}`);
                session.waitingForPrices = false;
                return;
            }

            // запись + фильтрация пустых строк
            productPrices = data.filter(row => row.length > 0 && row[0] !== '' && row[1] !== '');
            productPrices.shift();

            /**
             * Проверяет, что productPrices - это массив массивов,
             * где каждый вложенный массив содержит строку и число.
             * @param {Array} pricesArray - Массив цен.
             * @returns {boolean} - Возвращает true, если массив соответствует формату, иначе false.
             */
            const validateProductPrices = (pricesArray) => {
                // Проверяем, что pricesArray - это массив
                if (!Array.isArray(pricesArray)) {
                    return false;
                }

                // Проверяем каждый элемент массива
                for (const item of pricesArray) {
                    // Проверяем, что элемент - это массив
                    if (!Array.isArray(item)) {
                        return false;
                    }
                    // Проверяем, что массив содержит ровно два элемента
                    if (item.length !== 2) {
                        return false;
                    }
                    // Проверяем, что первый элемент - это строка, а второй - число
                    if (typeof item[0] !== 'string' || typeof item[1] !== 'number') {
                        return false;
                    }
                }
                
                return true;
            };

            if (!validateProductPrices(productPrices)) {
                console.error('Ошибка: productPrices не соответствует ожидаемому формату.');

                await ctx.reply('Таблицу которую вы скинули фуфло, скинь таблицу где данные начиная с A2+ являются названиями товаров (названия должны быть идентичны названиям из kaspi таблицы), а столбец B2+ только числа');
                session.waitingForPrices = false;
                return;
            }

            fs.unlinkSync(`uploads/${ctx.message.document.file_name}`);
            await ctx.reply('Запомнил товары и их цены');
            session.waitingForPrices = false;
        } else if (session.waitingForCount) {
            if (!productPrices || productPrices.length === 0) {
                await ctx.reply('Я не знаю себестоимость товаров. Воспользуйтесь командой /setprices');
                session.waitingForCount = false;
                return;
            }

            const fileId = ctx.message.document.file_id;
            const fileLink = await ctx.telegram.getFileLink(fileId);
            const response = await fetch(fileLink.href);
            const buffer = await response.buffer();

            fs.writeFileSync(`uploads/${ctx.message.document.file_name}`, buffer);

            await ctx.reply('Файл получен! Сейчас посчитаем...');
            
            const workbook = xlsx.readFile(`uploads/${ctx.message.document.file_name}`);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = xlsx.utils.sheet_to_json(sheet, {
                header: 1
            });

            if (data.length < 8) {
                await ctx.reply('Файл содержит недостаточно данных или неверный формат');
                fs.unlinkSync(`uploads/${ctx.message.document.file_name}`);
                session.waitingForCount = false;
                return;
            }

            const filteredData = data.slice(7).map(row => {
                return {
                    sumoper: row[18] || 0,
                    comm: row[20] || 0,
                    kaspicomm: row[26] || 0,
                    delivery: row[29] || 0,
                    product: row[31] || ''
                };
            });

            const findMatchingPrice = (productName, pricesArray) => {
                for (let [name, price] of pricesArray) {
                    if (productName.includes(name)) {
                        return price;
                    }
                }
                return null;
            }

            const updatedData = filteredData.map(item => {
                const price = findMatchingPrice(item.product, productPrices);
                return {
                    ...item,
                    price
                };
            });

            const headers = [
                "Наименование товара", 
                "Выгода", 
                "Цена продажи", 
                "Себестоимость", 
                "Коммиссия", 
                "Коммиссия каспи", 
                "Доставка",
                "Суммарная выгода"
            ];
            
            const dataForXLSX = updatedData.map((item, index) => [
                item.product,
                null,
                item.sumoper,
                item.price,
                item.comm,
                item.kaspicomm,
                item.delivery || 0,
                null
            ]);
            
            dataForXLSX.unshift(headers);
            
            const workbookNew = xlsx.utils.book_new();
            const worksheetNew = xlsx.utils.aoa_to_sheet(dataForXLSX);
            
            updatedData.forEach((item, idx) => {
                const rowIndex = idx + 2;
                worksheetNew[`B${rowIndex}`] = {
                    f: `C${rowIndex}+E${rowIndex}+F${rowIndex}+G${rowIndex}-D${rowIndex}`
                };
            });

            const totalRowIndex = dataForXLSX.length + 2;

            worksheetNew[`H2`] = {
                f: `SUM(B2:B${totalRowIndex - 1})`
            };

            const autoSizeColumns = (ws) => {
                const range = xlsx.utils.decode_range(ws['!ref']);
                const columnWidths = {};

                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const address = xlsx.utils.encode_cell({ c: C, r: R });
                        const cell = ws[address];
                        if (cell && cell.v) {
                            const value = cell.v.toString();
                            const length = value.length;

                            if (!columnWidths[C] || columnWidths[C] < length) {
                                columnWidths[C] = length;
                            }
                        }
                    }
                }

                ws['!cols'] = Object.keys(columnWidths).map(C => ({
                    width: columnWidths[C] + 2
                }));
            };

            autoSizeColumns(worksheetNew);

            xlsx.utils.book_append_sheet(workbookNew, worksheetNew, 'Результат');
            
            const resultBuffer = xlsx.write(workbookNew, { type: 'buffer', bookType: 'xlsx' });

            ctx.replyWithDocument({ source: resultBuffer, filename: 'result.xlsx' });

            fs.unlinkSync(`uploads/${ctx.message.document.file_name}`);

            session.waitingForCount = false;
        }
    } catch (error) {
        console.error('Ошибка обработки документа:', error);
        ctx.reply('Произошла ошибка при обработке вашего файла. Убедитесь, что он соответствует требованиям и попробуйте снова. \nПосле этого сообщения используйте команды из меню заново');
        session.waitingForCount = false;
        session.waitingForPrices = false;
    }
});

bot.launch();

// Запуск сервера
const PORT = process.env.PORT;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});