const XLSX = require('xlsx-style');
const fs = require('fs');
const path = require('path');
const { defineConfig } = require('cypress');

module.exports = defineConfig({
  e2e: {
    defaultCommandTimeout: 60000,
    pageLoadTimeout: 120000,
    requestTimeout: 30000,
    responseTimeout: 60000,
    setupNodeEvents(on, config) {
      on('task', {
        readExcel({ filePath }) {
          try {
            console.log('Reading Excel file from:', filePath);
            const absolutePath = path.resolve(filePath);
            console.log('Absolute path:', absolutePath);
            
            if (!fs.existsSync(absolutePath)) {
              console.log('❌ Excel file not found at:', absolutePath);
              return null;
            }
            
            console.log('Reading workbook...');
            const workbook = XLSX.readFile(absolutePath);
            console.log('Available sheets:', workbook.SheetNames);
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            console.log('Converting to JSON...');
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
              header: 1,
              raw: true,
              defval: ''
            });

            jsonData.slice(1).forEach(row => {
              if (Array.isArray(row)) {
                for (let i = 2; i <= 6; i++) {
                  if (row[i] !== undefined && row[i] !== '') {
                    row[i] = row[i].toString().replace('px', '').trim();
                  }
                }
                if (row[8] !== undefined && row[8] !== '') {
                  row[8] = row[8].toString().trim();
                }
              }
            });

            console.log('Excel Data Structure:');
            console.log('Number of rows:', jsonData.length);
            console.log('Headers:', jsonData[0]);
            console.log('First data row:', jsonData[1]);
            console.log('Second data row:', jsonData[2]);
            
            const headers = jsonData[0];
            headers.forEach((header, index) => {
              console.log(`Column ${index}: "${header}"`);
            });

            return jsonData;
          } catch (err) {
            console.error('Error reading Excel file:', err);
            return null;
          }
        },
        writeExcelSheets({ data, filename }) {
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

              Object.entries(data).forEach(([sheetName, sheetData]) => {
                if (!sheetData || !Array.isArray(sheetData) || sheetData.length === 0) {
                  console.log(`⚠️ No data for sheet "${sheetName}", skipping sheet creation.`);
                  return;
                }

                const worksheet = {};
                const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

                const columns = new Set();
                sheetData.forEach(row => {
                  if (row && typeof row === 'object') {
                    Object.keys(row).forEach(key => columns.add(key));
                  }
                });
                const columnArray = Array.from(columns);

                columnArray.forEach((col, idx) => {
                  const cellRef = XLSX.utils.encode_cell({ r: 0, c: idx });
                  worksheet[cellRef] = {
                    v: col,
                    t: 's',
                    s: {
                      font: { bold: false }
                    }
                  };
                  range.e.c = Math.max(range.e.c, idx);
                });

                sheetData.forEach((row, rowIdx) => {
                  if (!row || typeof row !== 'object') return;

                  columnArray.forEach((col, colIdx) => {
                    const cellRef = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
                    const value = row[col];

                    worksheet[cellRef] = {
                      v: value || '',
                      t: typeof value === 'number' ? 'n' : 's',
                      s: {}
                    };

                    if (col === 'Status' && value) {
                      if (value.includes('Mismatch')) {
                        worksheet[cellRef].s = {
                          fill: {
                            patternType: 'solid',
                            fgColor: { rgb: 'FF0000' }
                          },
                          font: {
                            bold: false,
                            color: { rgb: '#000000' }
                          }
                        };
                      } else if (value.includes('No Line Height in Guide')) {
                        worksheet[cellRef].s = {
                          fill: {
                            patternType: 'solid',
                            fgColor: { rgb: 'FFFF00' }
                          },
                          font: {
                            bold: false,
                            color: { rgb: '#000000' }
                          }
                        };
                      }

                      const mismatchTypes = {
                        'Font Size': ['Expected_fontSize', 'Actual_fontSize'],
                        'Font Weight': ['Expected_fontWeight', 'Actual_fontWeight'],
                        'Font Family': ['Expected_fontFamily', 'Actual_fontFamily'],
                        'Line Height': ['Expected_lineHeight', 'Actual_lineHeight'],
                      };

                      Object.entries(mismatchTypes).forEach(([type, cols]) => {
                        if (value.includes(type)) {
                          cols.forEach(colName => {
                            const colIdx = columnArray.indexOf(colName);
                            if (colIdx !== -1) {
                              const mismatchCellRef = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
                              if (!worksheet[mismatchCellRef]) {
                                worksheet[mismatchCellRef] = {
                                  v: row[colName] || '',
                                  t: 's',
                                  s: {}
                                };
                              }
                              worksheet[mismatchCellRef].s = {
                                fill: {
                                  patternType: 'solid',
                                  fgColor: { rgb: 'FF0000' }
                                },
                                font: {
                                  bold: false,
                                  color: { rgb: '#000000' }
                                }
                              };
                            }
                          });
                        }
                      });
                    }
                  });
                  range.e.r = Math.max(range.e.r, rowIdx + 1);
                });

                worksheet['!ref'] = XLSX.utils.encode_range(range);
                workbook.SheetNames.push(sheetName);
                workbook.Sheets[sheetName] = worksheet;
              });

              XLSX.writeFile(workbook, outputPath);
              console.log(`✅ Excel file with multiple sheets written to: ${outputPath}`);
              resolve(null);
            } catch (err) {
              console.error('❌ Error writing Excel sheets:', err);
              reject(err);
            }
          });
        }
      });

      return config;
    }
  }
});
