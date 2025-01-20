import fs from "fs";
import ExcelJS from "exceljs";
import XlsxTemplate from "xlsx-template";
import libre from "libreoffice-convert";
import path from "path";
import { parse, getWeek } from "date-fns";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import {
  STOPS_PER_HOUR,
  START_INVOICE_NR,
  ADMINISTRATIONAL_COST,
  STOP_PRICE,
  INVOICE_NUMBER_FORMAT,
} from "./constants.js";
import { fileURLToPath } from "url";

// Resolve the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Parses a PDF file to extract relevant data.
 * @param {string} filePath - The path to the PDF file.
 * @returns {Promise<Object>} - The extracted data including date, total stops, week of year, and converted hours.
 * @throws {Error} - Throws an error if required data is not found in the PDF.
 */
export async function parsePDF(filePath) {
  const dataBuffer = fs.readFileSync(filePath);
  const dataArray = new Uint8Array(dataBuffer); // Convert Buffer to Uint8Array
  const standardFontDataUrl = `${path.join(
    __dirname,
    "node_modules",
    "pdfjs-dist",
    "standard_fonts"
  )}/`;
  const pdfDoc = await getDocument({ data: dataArray, standardFontDataUrl })
    .promise;
  const text = await extractTextFromPDF(pdfDoc);

  const dateMatch = text.match(
    /Activiteitenrapport\s*([0-9]{2}-[0-9]{2}-[0-9]{4})/
  );
  const stopsMatch = text.match(/Totaal aantal succesvolle stops\s*([0-9]+)/);

  if (dateMatch && stopsMatch) {
    const date = dateMatch[1];
    const totalStops = parseInt(stopsMatch[1], 10);
    const dateObj = parse(date, "dd-MM-yyyy", new Date());
    const weekOfYear = getWeek(dateObj);
    const convertedHours =
      Math.round((totalStops / STOPS_PER_HOUR) * 100) / 100;

    return {
      date,
      totalStops,
      weekOfYear,
      convertedHours,
    };
  } else {
    throw new Error(`Required data not found in PDF: ${filePath}`);
  }
}

/**
 * Extracts text content from a PDF document.
 * @param {Object} pdfDoc - The PDF document object.
 * @returns {Promise<string>} - The extracted text content.
 */
async function extractTextFromPDF(pdfDoc) {
  let text = "";
  const numPages = pdfDoc.numPages;
  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const content = await page.getTextContent();
    const pageText = content.items.map((item) => item.str).join(" ");
    text += pageText + "\n";
  }
  return text;
}

/**
 * Gets the column letters from the header row of an Excel worksheet.
 * @param {Object} worksheet - The Excel worksheet object.
 * @returns {Object} - An object mapping column names to their respective letters.
 */
export function getColumnLetters(worksheet) {
  const headerRow = worksheet.getRow(1); // Assuming the first row contains headers
  const columnLetters = {};

  headerRow.eachCell((cell) => {
    columnLetters[cell.text.trim().toLowerCase()] = cell.address.replace(
      /[0-9]/g,
      ""
    );
  });

  return columnLetters;
}

/**
 * Extracts the numeric part from an invoice number string.
 * @param {string} invoiceNumber - The invoice number string (e.g., 'INVOICE #001').
 * @returns {Array|null} - An array containing the full match and the captured group, or null if no match is found.
 */
export function getNumberFromInvoice(invoiceNumber) {
  const matchRegex = new RegExp(
    INVOICE_NUMBER_FORMAT.replace("{number}", "(\\d+)")
  );
  const number = invoiceNumber.match(matchRegex);
  return number;
}

/**
 * Gets the next invoice number based on the last invoice number.
 * @param {string} lastInvoiceNumber - The last invoice number.
 * @returns {number} - The next invoice number.
 */
export function getNextNumber(lastInvoiceNumber) {
  const match = getNumberFromInvoice(lastInvoiceNumber);
  if (match) {
    const number = parseInt(match[1], 10) + 1;
    return number;
  }
  return 1;
}

/**
 * Generates the next invoice number string.
 * @param {string} lastInvoiceNumber - The last invoice number.
 * @returns {string} - The next invoice number string.
 */
export function getNextInvoiceNumber(lastInvoiceNumber) {
  const number = getNextNumber(lastInvoiceNumber);
  return INVOICE_NUMBER_FORMAT.replace(
    "{number}",
    String(number).padStart(3, "0")
  );
}

/**
 * Updates an Excel file with new data and returns uninvoiced data and the next invoice number.
 * @param {string} filePath - The path to the Excel file.
 * @param {Array<Object>} data - The data to update the Excel file with.
 * @returns {Promise<Object>} - An object containing uninvoiced data and the next invoice number.
 */
export async function updateExcel(filePath, data) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);
  const columnLetters = getColumnLetters(worksheet);
  const existingDates = new Set();
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    // Skip header row
    if (rowNumber > 1) {
      const dateCell = row.getCell(columnLetters["date"]);
      if (dateCell.value) {
        existingDates.add(dateCell.value);
      }
    }
  });

  let lastInvoiceNr = START_INVOICE_NR;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > 1) {
      const invoiceNrCell = row.getCell(columnLetters["invoice.nr"]);
      if (invoiceNrCell.value) {
        console.log(
          "Found a previous invoice in reports file",
          invoiceNrCell.value
        );
        lastInvoiceNr = invoiceNrCell.value;
      }
    }
  });

  const nextInvoiceNumber = getNextInvoiceNumber(lastInvoiceNr);
  console.log("Next invoice number", nextInvoiceNumber);

  data.forEach(({ date, totalStops, weekOfYear, convertedHours }) => {
    if (!existingDates.has(date)) {
      const newRow = worksheet.lastRow.number + 1;
      worksheet.getCell(`${columnLetters["date"]}${newRow}`).value = date;
      worksheet.getCell(`${columnLetters["total stops"]}${newRow}`).value =
        totalStops;
      worksheet.getCell(`${columnLetters["week of year"]}${newRow}`).value =
        weekOfYear;
      worksheet.getCell(`${columnLetters["converted hours"]}${newRow}`).value =
        convertedHours;
      worksheet.getCell(`${columnLetters["invoice.nr"]}${newRow}`).value =
        nextInvoiceNumber;
    }
  });

  await workbook.xlsx.writeFile(filePath);

  // Collect uninvoiced data grouped by week
  const uninvoicedData = {};
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      const invoicedCell = row.getCell(columnLetters["invoiced"]);

      if (invoicedCell.value !== true) {
        const date = row.getCell(columnLetters["date"]).value;
        const totalStops = row.getCell(columnLetters["total stops"]).value;
        const weekOfYear = row.getCell(columnLetters["week of year"]).value;
        const convertedHours = row.getCell(
          columnLetters["converted hours"]
        ).value;

        if (!uninvoicedData[weekOfYear]) {
          uninvoicedData[weekOfYear] = [];
        }
        uninvoicedData[weekOfYear].push({
          date,
          totalStops,
          weekOfYear,
          convertedHours,
        });
      }
    }
  });

  return { uninvoicedData, nextInvoiceNumber };
}

/**
 * Converts an Excel file to a PDF file.
 * @param {string} inputPath - The path to the input Excel file.
 * @param {string} outputPath - The path to the output PDF file.
 * @returns {Promise<string>} - A promise that resolves with a success message.
 * @throws {Error} - Throws an error if the conversion fails.
 */
export async function convertExcelToPdf(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    const file = fs.readFileSync(inputPath);
    libre.convert(file, ".pdf", undefined, (err, done) => {
      if (err) {
        reject(`Error converting file: ${err.message}`);
      } else {
        fs.writeFileSync(outputPath, done);
        resolve(`Successfully converted ${inputPath} to ${outputPath}`);
      }
    });
  });
}

/**
 * Creates a PDF invoice based on an Excel template.
 * @param {string} templatePath - The path to the Excel template file.
 * @param {string} outputPath - The path to the output PDF file.
 * @param {Object} data - The data to populate the template with.
 * @param {string} invoiceNumber - The invoice number to use in the template.
 * @returns {Promise<void>} - A promise that resolves when the PDF is created.
 * @throws {Error} - Throws an error if the PDF creation or conversion fails.
 */
export async function createPDF(templatePath, outputPath, data, invoiceNumber) {
  const templateData = fs.readFileSync(templatePath);
  const template = new XlsxTemplate(templateData);
  const sheetNumber = 1;

  // Combine all data into a format suitable for the template
  const weeks = Object.keys(data);
  const records = [];
  weeks.forEach((weekOfYear, index) => {
    const weekData = data[weekOfYear];
    const totalStops = weekData.reduce(
      (sumStops, { totalStops }) => sumStops + totalStops,
      0
    );
    const totalHours = totalStops / STOPS_PER_HOUR;
    const ADMINISTRATIONAL_HOURS =
      ADMINISTRATIONAL_COST / (STOPS_PER_HOUR * STOP_PRICE);
    const hours =
      Math.round(
        (index === 0 ? totalHours - ADMINISTRATIONAL_HOURS : totalHours) * 100
      ) / 100;
    const total = hours * STOPS_PER_HOUR * STOP_PRICE;
    records.push({
      weekOfYear,
      hours,
      price: STOPS_PER_HOUR,
      total,
    });
  });

  // Substitute the table and the single cell placeholders in the template
  template.substitute(sheetNumber, {
    invoiceNumber,
    records,
  });

  // Generate the XLSX from the template
  const xlsxBuffer = template.generate({ type: "nodebuffer" });

  // Save the buffer to a temporary Excel file
  const tempExcelPath = outputPath.replace(".pdf", ".xlsx");
  fs.writeFileSync(tempExcelPath, xlsxBuffer);
  // Validate the generated XLSX file
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempExcelPath);
    console.log("Temporary Excel file is valid.");
  } catch (error) {
    throw new Error(`Error validating temporary Excel file: ${error.message}`);
  }
  // Use libreoffice-convert to convert the Excel file to PDF
  await convertExcelToPdf(tempExcelPath, outputPath);

  // Clean up the temporary Excel file
  fs.unlinkSync(tempExcelPath);
}

/**
 * Marks data entries as invoiced in an Excel file.
 * @param {string} filePath - The path to the Excel file.
 * @param {Object} data - The data entries to mark as invoiced.
 * @returns {Promise<void>} - A promise that resolves when the entries are marked.
 */
export async function markAsInvoiced(filePath, data) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.getWorksheet(1);
  const columnLetters = getColumnLetters(worksheet);

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip header row
      const dateCell = row.getCell(columnLetters["date"]);
      if (
        Object.values(data)
          .flat()
          .some((entry) => entry.date === dateCell.value)
      ) {
        row.getCell(columnLetters["invoiced"]).value = true;
      }
    }
  });

  await workbook.xlsx.writeFile(filePath);
}

/**
 * Recursively get all PDF files from a directory and its subdirectories
 * @param {string} directory - The directory to search
 * @returns {string[]} - Array of PDF file paths
 */
export function getPdfFiles(directory) {
  let results = [];
  const files = fs.readdirSync(directory);

  files.forEach((file) => {
    const filePath = path.join(directory, file);
    const stat = fs.statSync(filePath);

    if (stat && stat.isDirectory()) {
      // Recurse into subdirectory
      results = results.concat(getPdfFiles(filePath));
    } else if (path.extname(file).toLowerCase() === ".pdf") {
      // Add PDF file to results
      results.push(filePath);
    }
  });

  return results;
}
