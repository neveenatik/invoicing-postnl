import fs from 'fs';
import ExcelJS from 'exceljs';
import XlsxTemplate from 'xlsx-template';
import libre from 'libreoffice-convert';
import path from 'path';
import { parse, getWeek } from 'date-fns';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { STOPS_PER_HOUR, START_INVOICE_NR } from './constants.js';
import { fileURLToPath } from 'url';

// Resolve the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export async function parsePDF(filePath) {
    const dataBuffer = fs.readFileSync(filePath);
    const dataArray = new Uint8Array(dataBuffer); // Convert Buffer to Uint8Array
    const standardFontDataUrl = `${path.join(__dirname, 'node_modules', 'pdfjs-dist', 'standard_fonts')}/`;
    const pdfDoc = await getDocument({ data: dataArray, standardFontDataUrl }).promise;
    const text = await extractTextFromPDF(pdfDoc);

    const dateMatch = text.match(/Activiteitenrapport\s*([0-9]{2}-[0-9]{2}-[0-9]{4})/);
    const stopsMatch = text.match(/Totaal aantal succesvolle stops\s*([0-9]+)/);

    if (dateMatch && stopsMatch) {
        const date = dateMatch[1];
        const totalStops = parseInt(stopsMatch[1], 10);
        const dateObj = parse(date, 'dd-MM-yyyy', new Date());
        const weekOfYear = getWeek(dateObj);
        const convertedHours = totalStops / STOPS_PER_HOUR; // Calculate and round to two decimals

        return {
            date,
            totalStops,
            weekOfYear,
            convertedHours
        };
    } else {
        throw new Error(`Required data not found in PDF: ${filePath}`);
    }
}

async function extractTextFromPDF(pdfDoc) {
    let text = '';
    const numPages = pdfDoc.numPages;
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
        const page = await pdfDoc.getPage(pageNum);
        const content = await page.getTextContent();
        const pageText = content.items.map(item => item.str).join(' ');
        text += pageText + '\n';
    }
    return text;
}

export function getColumnLetters(worksheet) {
    const headerRow = worksheet.getRow(1); // Assuming the first row contains headers
    const columnLetters = {};

    headerRow.eachCell((cell) => {
        columnLetters[cell.text.trim().toLowerCase()] = cell.address.replace(/[0-9]/g, '');
    });

    return columnLetters;
}

export function getNextInvoiceNumber(lastInvoiceNumber) {
    const match = lastInvoiceNumber.match(/INVOICE #(\d+)/);
    if (match) {
        const number = parseInt(match[1], 10) + 1;
        return `INVOICE #${String(number).padStart(3, '0')}`;
    }
    return 'INVOICE #001';
}

export async function updateExcel(filePath, data) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);
    const columnLetters = getColumnLetters(worksheet);
    const existingDates = new Set();
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) { // Skip header row
            const dateCell = row.getCell(columnLetters['date']);
            if (dateCell.value) {
                existingDates.add(dateCell.value);
            }
        }
    });

    let lastInvoiceNr = START_INVOICE_NR;
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) {
            const invoiceNrCell = row.getCell(columnLetters['invoice.nr']);
            if (invoiceNrCell.value) {
                lastInvoiceNr = invoiceNrCell.value;
            }
        }
    });

    const nextInvoiceNumber = getNextInvoiceNumber(lastInvoiceNr);

    data.forEach(({ date, totalStops, weekOfYear, convertedHours }) => {
        if (!existingDates.has(date)) {
            const newRow = worksheet.lastRow.number + 1;
            worksheet.getCell(`${columnLetters['date']}${newRow}`).value = date;
            worksheet.getCell(`${columnLetters['total stops']}${newRow}`).value = totalStops;
            worksheet.getCell(`${columnLetters['week of year']}${newRow}`).value = weekOfYear;
            worksheet.getCell(`${columnLetters['converted hours']}${newRow}`).value = convertedHours;
            worksheet.getCell(`${columnLetters['invoice.nr']}${newRow}`).value = nextInvoiceNumber;
        }
    });

    await workbook.xlsx.writeFile(filePath);

    // Collect uninvoiced data grouped by week
    const uninvoicedData = {};
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) { // Skip header row
            const invoicedCell = row.getCell(columnLetters['invoiced']);

            if (invoicedCell.value !== true) {
                const date = row.getCell(columnLetters['date']).value;
                const totalStops = row.getCell(columnLetters['total stops']).value;
                const weekOfYear = row.getCell(columnLetters['week of year']).value;
                const convertedHours = row.getCell(columnLetters['converted hours']).value;

                if (!uninvoicedData[weekOfYear]) {
                    uninvoicedData[weekOfYear] = [];
                }
                uninvoicedData[weekOfYear].push({ date, totalStops, weekOfYear, convertedHours });
            }
        }
    });

    return { uninvoicedData, nextInvoiceNumber };
}

export async function convertExcelToPdf(inputPath, outputPath) {
    return new Promise((resolve, reject) => {
        const file = fs.readFileSync(inputPath);
        libre.convert(file, '.pdf', undefined, (err, done) => {
            if (err) {
                reject(`Error converting file: ${err.message}`);
            } else {
                fs.writeFileSync(outputPath, done);
                resolve(`Successfully converted ${inputPath} to ${outputPath}`);
            }
        });
    });
}

export async function createPDF(templatePath, outputPath, data, invoiceNumber) {
    const templateData = fs.readFileSync(templatePath);
    const template = new XlsxTemplate(templateData);
    const sheetNumber = 1;
    // Combine all data into a format suitable for the template
    const weeks = Object.keys(data);
    const records = [];
    weeks.forEach(weekOfYear => {
        const weekData = data[weekOfYear];
        const totalHours = weekData.reduce((sumHours, { convertedHours }) => sumHours + parseFloat(convertedHours), 0);
        records.push({
            weekOfYear, hours: totalHours // Round hours to two decimals
        });
    });
    template.substitute(sheetNumber, {
        invoiceNumber,
        records
    });

    // Generate the XLSX from the template
    const xlsxBuffer = template.generate({ type: 'nodebuffer' });

    // Save the buffer to a temporary Excel file
    const tempExcelPath = outputPath.replace('.pdf', '.xlsx');
    fs.writeFileSync(tempExcelPath, xlsxBuffer);
    // Validate the generated XLSX file
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(tempExcelPath);
        console.log('Temporary Excel file is valid.');
    } catch (error) {
        throw new Error(`Error validating temporary Excel file: ${error.message}`);
    }
    // Use libreoffice-convert to convert the Excel file to PDF
    await convertExcelToPdf(tempExcelPath, outputPath);

    // Clean up the temporary Excel file
    fs.unlinkSync(tempExcelPath);
}

export async function markAsInvoiced(filePath, data) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(1);
    const columnLetters = getColumnLetters(worksheet);

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) { // Skip header row
            const dateCell = row.getCell(columnLetters['date']);
            if (Object.values(data).flat().some(entry => entry.date === dateCell.value)) {
                row.getCell(columnLetters['invoiced']).value = true;
            }
        }
    });

    await workbook.xlsx.writeFile(filePath);
}

export function getPdfFiles(directory) {
    const files = fs.readdirSync(directory);
    return files.filter(file => path.extname(file).toLowerCase() === '.pdf').map(file => path.join(directory, file));
}
