const XLSX = require('xlsx-style');
const fs = require('fs');
const path = require('path');
const { defineConfig } = require('cypress');

module.exports = defineConfig({
  e2e: {
    baseUrl: 'https://www.rehabclinics.com',
    setupNodeEvents(on, config) {
      on('task', {
        readExcel({ filePath }) {
          try {
            const absolutePath = path.resolve(filePath);
            if (!fs.existsSync(absolutePath)) {
              return null;
            }
            const workbook = XLSX.readFile(absolutePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            return jsonData;
          } catch (err) {
            console.error('Error reading Excel file:', err);
            return null;
          }
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
              const workbook = {
                SheetNames: [],
                Sheets: {}
              };

              // For each viewport, add a sheet with its data
              Object.entries(data).forEach(([sheetName, sheetData]) => {
                if (!sheetData || !Array.isArray(sheetData) || sheetData.length === 0) {
                  console.log(`⚠️ No data for sheet "${sheetName}", skipping sheet creation.`);
                  return;
                }

                // Create worksheet manually
                const worksheet = {};
                const range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

                // Get all possible columns from the data
                const columns = new Set();
                sheetData.forEach(row => {
                  if (row && typeof row === 'object') {
                    Object.keys(row).forEach(key => columns.add(key));
                  }
                });
                const columnArray = Array.from(columns);

                // Add header row
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

                // Add data rows
                sheetData.forEach((row, rowIdx) => {
                  if (!row || typeof row !== 'object') return;

                  columnArray.forEach((col, colIdx) => {
                    const cellRef = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
                    const value = row[col];

                    // Initialize cell with value and type
                    worksheet[cellRef] = {
                      v: value || '',
                      t: typeof value === 'number' ? 'n' : 's',
                      s: {} // Initialize style object
                    };

                    // Apply styling for mismatches
                    if (col === 'Status' && value && value.includes('Mismatch')) {
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

                      // Style the mismatched value cells
                      const mismatchTypes = {
                        'Font Size': ['Expected_fontSize', 'Actual_fontSize'],
                        'Font Weight': ['Expected_fontWeight', 'Actual_fontWeight'],
                        'Font Family': ['Expected_fontFamily', 'Actual_fontFamily']
                      };

                      Object.entries(mismatchTypes).forEach(([type, cols]) => {
                        if (value.includes(type)) {
                          cols.forEach(colName => {
                            const colIdx = columnArray.indexOf(colName);
                            if (colIdx !== -1) {
                              const mismatchCellRef = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
                              // Ensure the cell exists and has a style object
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

                // Set worksheet range
                worksheet['!ref'] = XLSX.utils.encode_range(range);

                // Add the worksheet to the workbook
                workbook.SheetNames.push(sheetName);
                workbook.Sheets[sheetName] = worksheet;
              });

              // Write to file
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
});
