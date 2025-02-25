# Invoicing Script

This script processes PDF files, updates data in an Excel file, generates invoices based on a template, and marks the entries as invoiced. Below are the instructions on how to use the script.

## Prerequisites

Ensure you have the following installed:
- Node.js
- npm (Node Package Manager)
- LibreOffice (required for `libreoffice-convert` library)

## Installation

1. Clone the repository or download the script to your local machine.
2. Navigate to the directory containing the script in your terminal or Command Prompt for windows users.
2. Open Command Prompt (for Windows users) or terminal (for MacOS users) and navigate to the directory containing the script.
3. Install the necessary dependencies by running:
   ```cmd
   npm install
   ```

### Installing LibreOffice

LibreOffice is required for converting Excel files to PDF. You can download and install LibreOffice from [here](https://www.libreoffice.org/download/download/).

After installing LibreOffice, ensure that the `soffice` executable is in your system's PATH. You can verify this by running the following command in Command Prompt:

```cmd
soffice --version
```

If the command returns the LibreOffice version, it means LibreOffice is correctly installed and in your PATH.

## Usage

To run the script, use the following command:

```bash
node invoicing.js <excelPath> <templatePath> <outputDirectory> <pdfDirectory>
```

### Arguments

- `<excelPath>`: Path to the Excel file where the data will be updated.
- `<templatePath>`: Path to the Excel template file used for generating invoices.
- `<outputDirectory>`: Directory where the generated invoice PDF will be saved.
- `<pdfDirectory>`: Directory containing the PDF files to be processed.

### Example

MacOS:
```bash
node invoicing.js path/to/your/excel.xlsx path/to/your/template.xlsx path/to/output/directory path/to/pdf/directory
```

Windows:
```cmd
node invoicing.js C:\path\to\your\excel.xlsx C:\path\to\your\template.xlsx C:\path\to\output\directory C:\path\to\pdf\directory
```

### Description

1. The script will resolve and validate the paths provided.
2. It will fetch all PDF files from the specified `pdfDirectory`.
3. Each PDF file will be parsed to extract data.
4. The extracted data will be used to update the Excel file specified by `excelPath`.
5. If there are uninvoiced entries, a new invoice PDF will be generated using the `templatePath` and saved in the `outputDirectory`.
6. The entries in the Excel file will be marked as invoiced.

### Output

- The script will log the paths being used and the progress of each step.
- If there are new entries to process, a new invoice PDF will be generated and saved in the specified output directory.
- If there are no new entries to process, it will log "No new entries to process. No invoice generated."

### Template Example
The repository includes a template example file (MTNA_invoice.xlsx) that demonstrates the values that can be used in the provided template. This file serves as a guide for how to structure your template to work with the script.

### Report Excel File
The repository also includes a report Excel file (report.xlsx) that provides an example of what data would be added to the report file. This file shows the expected format and structure of the data entries.

### Error Handling

If an error occurs during execution, it will be logged to the console.

## Contributors

- [neveenatik]

## License

This project is licensed under the MIT License.

---
