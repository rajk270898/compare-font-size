const Excel = require('exceljs');
const fs = require('fs');
const path = require('path');

module.exports = (on, config) => {
  on('task', {
    readTypographyData(filePath) {
      const workbook = new Excel.Workbook();
      return workbook.xlsx.readFile(filePath).then(() => {
        const worksheet = workbook.getWorksheet('Typography');
        const data = [];
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row
          data.push({
            selector: row.getCell(1).text,
            property: row.getCell(2).text,
            expectedValue: row.getCell(3).text,
            viewport: row.getCell(4).text,
          });
        });
        return data;
      });
    },

    async writeResultToExcel({ selector, property, expectedValue, actual, viewport, status }) {
      const filePath = path.resolve('cypress/results/typography_results.xlsx');
      const workbook = new Excel.Workbook();
      let worksheet;
      
      if (fs.existsSync(filePath)) {
        // If file exists, load it
        await workbook.xlsx.readFile(filePath);
        worksheet = workbook.getWorksheet('Results');
      }
      
      // If worksheet doesn't exist, create and add headers
      if (!worksheet) {
        worksheet = workbook.addWorksheet('Results');
        worksheet.addRow(['Selector', 'Property', 'Expected', 'Actual', 'Viewport', 'Status']);
      }

      // Add the result row
      worksheet.addRow([selector, property, expectedValue, actual, viewport, status]);

      // Save the workbook
      await workbook.xlsx.writeFile(filePath);

      return true;
    },
  });
};
