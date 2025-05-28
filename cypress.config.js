const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

module.exports = {
  e2e: {
    setupNodeEvents(on, config) {
      on('task', {
        readExcel({ filePath }) {
          return new Promise((resolve, reject) => {
            try {
              const absolutePath = path.resolve(filePath);
              const workbook = XLSX.readFile(absolutePath);
              const sheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[sheetName];
              const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
              resolve(jsonData);
            } catch (err) {
              console.error('Error reading Excel file:', err);
              reject(err);
            }
          });
        },

        writeExcelSheets({ data, filename }) {
          return new Promise((resolve, reject) => {
            try {
              // Ensure directory exists
              const outputPath = path.resolve(filename);
              const dirPath = path.dirname(outputPath);

              if (!fs.existsSync(dirPath)) {
                fs.mkdirSync(dirPath, { recursive: true });
              }

              // Create new workbook
              const workbook = XLSX.utils.book_new();

              // For each viewport, add a sheet with its data
              Object.entries(data).forEach(([sheetName, sheetData]) => {
                if (!sheetData || sheetData.length === 0) {
                  console.log(`⚠️ No data for sheet "${sheetName}", skipping sheet creation.`);
                  return;
                }

                const worksheet = XLSX.utils.json_to_sheet(sheetData);
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
              });

              // Write workbook to file
              XLSX.writeFile(workbook, outputPath);

              console.log(`✅ Excel file with multiple sheets written to: ${outputPath}`);

              resolve(null);
            } catch (err) {
              console.error('❌ Error writing Excel sheets:', err);
              reject(err);
            }
          });
        },
      });
    },
  },
};
