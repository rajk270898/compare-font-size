const XLSX = require('xlsx');
const fs = require('fs');

module.exports = (on, config) => {
  on('task', {
    readExcel({ filePath }) {
      const workbook = XLSX.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      return XLSX.utils.sheet_to_json(sheet, { header: 1 });
    },

    writeExcelSheets({ data, filename }) {
      const workbook = XLSX.utils.book_new();

      Object.keys(data).forEach(sheetName => {
        const sheetData = data[sheetName];
        const worksheet = XLSX.utils.json_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      });

      XLSX.writeFile(workbook, `cypress/results/${filename}`);
      return null;
    },
  });
};
