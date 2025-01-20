import { parsePDF, updateExcel, createPDF, markAsInvoiced, getPdfFiles } from './helpers.js';
import { INVOICE_FILE_NAME_FORMAT } from './constants.js'
import path from 'path';
import { fileURLToPath } from 'url';

// Resolve the directory name
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

(async () => {
    try {
        const args = process.argv.slice(2);
        if (args.length < 4) {
            console.log('Usage: node invoicing.js <excelPath> <templatePath> <outputDirectory> <pdfDirectory>');
            console.log('Example: node invoicing.js path/to/your/excel.xlsx path/to/your/template.xlsx path/to/output/directory path/to/pdf/directory');
            process.exit(1);
        }

        const [excelPath, templatePath, outputDirectory, pdfDirectory] = args.map(arg => path.resolve(__dirname, arg));

        console.log(`Using excelPath: ${excelPath}`);
        console.log(`Using templatePath: ${templatePath}`);
        console.log(`Using outputDirectory: ${outputDirectory}`);
        console.log(`Using pdfDirectory: ${pdfDirectory}`);

        const pdfPaths = getPdfFiles(pdfDirectory);
        console.log('PDF paths:', pdfPaths);
        
        const allDataPromises = pdfPaths.map(pdfPath => parsePDF(pdfPath));
        const allData = await Promise.all(allDataPromises);

        const { uninvoicedData, nextInvoiceNumber } = await updateExcel(excelPath, allData);

        if (Object.keys(uninvoicedData).length > 0) {
            const outputFilePath = path.join(outputDirectory, `${INVOICE_FILE_NAME_FORMAT.replace('{number}', nextInvoiceNumber)}.pdf`);
            await createPDF(templatePath, outputFilePath, uninvoicedData, nextInvoiceNumber);
            await markAsInvoiced(excelPath, uninvoicedData);
            console.log('Process completed successfully!');
        } else {
            console.log('No new entries to process. No invoice generated.');
        }
    } catch (error) {
        console.error('Error:', error);
    }
})();
