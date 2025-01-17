# Invoicing Script

This script processes PDF files, updates data in an Excel file, generates invoices based on a template, and marks the entries as invoiced. Below are the instructions on how to use the script.

## Prerequisites

Ensure you have the following installed:
- Node.js
- npm (Node Package Manager)

## Installation

1. Clone the repository or download the script to your local machine.
2. Navigate to the directory containing the script in your terminal.
3. Install the necessary dependencies by running:
   ```bash
   npm install
   ```

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

```bash
node invoicing.js path/to/your/excel.xlsx path/to/your/template.xlsx path/to/output/directory path/to/pdf/directory
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

### Error Handling

If an error occurs during execution, it will be logged to the console.

## Contributors

- [Your Name]

## License

This project is licensed under the MIT License.

---
