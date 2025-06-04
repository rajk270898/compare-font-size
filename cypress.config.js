const XLSX = require('xlsx-style');
const fs = require('fs');
const path = require('path');
const { defineConfig } = require('cypress');

module.exports = defineConfig({
  e2e: {
    setupNodeEvents(on, config) {
      on('task', {
        readExcel({ filePath }) {
          try {
            const absolutePath = path.resolve(filePath);
            if (!fs.existsSync(absolutePath)) {
              console.log('❌ Excel file not found at:', absolutePath);
              return null;
            }

            const workbook = XLSX.readFile(absolutePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              raw: true,
              defval: ''
            });

            return jsonData;
          } catch (err) {
            console.error('❌ Error reading Excel file:', err);
            return null;
          }
        },

        writeExcelSheets({ data, filename, sheetOrder }) {
          return new Promise((resolve, reject) => {
            try {
              const outputPath = path.resolve(filename);
              const dirPath = path.dirname(outputPath);

              if (!fs.existsSync(dirPath)) {
                fs.mkdirSync(dirPath, { recursive: true });
              }

              const workbook = {
                SheetNames: [],
                Sheets: {}
              };

              (sheetOrder || Object.keys(data)).forEach(sheetName => {
                const sheetData = data[sheetName];
                if (!sheetData || !Array.isArray(sheetData) || sheetData.length === 0) return;

                const worksheet = {};
                const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
                const columns = new Set();

                sheetData.forEach(row => {
                  if (row && typeof row === 'object') {
                    Object.keys(row).forEach(key => columns.add(key));
                  }
                });

                const columnArray = Array.from(columns);

                // Write header row
                columnArray.forEach((col, idx) => {
                  const cellRef = XLSX.utils.encode_cell({ r: 0, c: idx });
                  worksheet[cellRef] = {
                    v: col,
                    t: 's',
                    s: { font: { bold: true } }
                  };
                  range.e.c = Math.max(range.e.c, idx);
                });

                sheetData.forEach((row, rowIdx) => {
                  columnArray.forEach((col, colIdx) => {
                    const cellRef = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
                    const value = row[col];

                    worksheet[cellRef] = {
                      v: value || '',
                      t: typeof value === 'number' ? 'n' : 's',
                      s: {}
                    };

                    // Highlight mismatched values
                    if (col === 'Status' && value && value.includes('Mismatch')) {
                      worksheet[cellRef].s = {
                        fill: {
                          patternType: 'solid',
                          fgColor: { rgb: 'FF0000' }
                        },
                        font: {
                          color: { rgb: '000000' }
                        }
                      };

                      const mismatchTypes = {
                        'Font Size': ['Expected_fontSize', 'Actual_fontSize'],
                        'Font Weight': ['Expected_fontWeight', 'Actual_fontWeight'],
                        'Font Family': ['Expected_fontFamily', 'Actual_fontFamily'],
                        'Line Height': ['Expected_lineHeight', 'Actual_lineHeight']
                      };

                      Object.entries(mismatchTypes).forEach(([type, cols]) => {
                        if (value.includes(type)) {
                          const isFontWeight = type === 'Font Weight';
                          const expectedCol = cols[0];
                          const expectedVal = row[expectedCol];

                          // Skip Font Weight mismatch if Expected value is missing or '-'
                          if (!isFontWeight || (expectedVal && expectedVal !== '-')) {
                            cols.forEach(colName => {
                              const colIndex = columnArray.indexOf(colName);
                              if (colIndex !== -1) {
                                const mismatchCell = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIndex });
                                worksheet[mismatchCell].s = {
                                  fill: {
                                    patternType: 'solid',
                                    fgColor: { rgb: 'FF0000' }
                                  },
                                  font: {
                                    color: { rgb: '000000' }
                                  }
                                };
                              }
                            });
                          }
                        }
                      });
                    }

                    range.e.r = Math.max(range.e.r, rowIdx + 1);
                  });
                });

                worksheet['!ref'] = XLSX.utils.encode_range(range);
                workbook.SheetNames.push(sheetName);
                workbook.Sheets[sheetName] = worksheet;
              });

              XLSX.writeFile(workbook, outputPath);
              console.log(`✅ Excel written to: ${outputPath}`);
              resolve(null);
            } catch (err) {
              console.error('❌ Error writing Excel file:', err);
              reject(err);
            }
          });
        }
      });

      return config;
    }
  }
});
