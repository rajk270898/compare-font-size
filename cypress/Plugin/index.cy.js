const Excel = require('exceljs');

module.exports = (on, config) => {
  on('task', {
    readTypographyData(filePath) {
      const workbook = new Excel.Workbook();
      return workbook.xlsx.readFile(filePath).then(() => {
        const worksheet = workbook.getWorksheet('Typography');
        const data = [];
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header
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

    writeResultToExcel({ selector, property, expectedValue, actual, viewport, status }) {
      const filePath = 'cypress/results/typography_results.xlsx';
      const workbook = fs.existsSync(filePath) ? new Excel.Workbook() : new Excel.Workbook();
      const sheetName = 'Results';

      return workbook.xlsx.readFile(filePath).then(() => {
        let worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          worksheet = workbook.addWorksheet(sheetName);
          worksheet.addRow(['Selector', 'Property', 'Expected', 'Actual', 'Viewport', 'Status']);
        }
        worksheet.addRow([selector, property, expectedValue, actual, viewport, status]);
        return workbook.xlsx.writeFile(filePath).then(() => true);
      });
    },
  });
};
