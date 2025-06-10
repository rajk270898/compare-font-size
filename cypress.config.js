const XLSX = require('xlsx-style');
const Excel = require('exceljs');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');
const { defineConfig } = require('cypress');

module.exports = defineConfig({
  e2e: {
    setupNodeEvents(on, config) {
      on('task', {

        // ✅ Read style guide data
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

        // ✅ Write styled result sheets
        writeExcelSheets({ data, filename, sheetOrder }) {
          return new Promise((resolve, reject) => {
            try {
              const outputPath = path.resolve(filename);
              const dirPath = path.dirname(outputPath);

              if (!fs.existsSync(dirPath)) {
                fs.mkdirSync(dirPath, { recursive: true });
              }

              const workbook = { SheetNames: [], Sheets: {} };

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

                // Header row
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

                    // Highlight mismatches
                    if (col === 'Status' && value && value.includes('Mismatch')) {
                      worksheet[cellRef].s = {
                        fill: { patternType: 'solid', fgColor: { rgb: 'FF0000' } },
                        font: { color: { rgb: '000000' } }
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

                          if (!isFontWeight || (expectedVal && expectedVal !== '-')) {
                            cols.forEach(colName => {
                              const colIndex = columnArray.indexOf(colName);
                              if (colIndex !== -1) {
                                const mismatchCell = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIndex });
                                worksheet[mismatchCell].s = {
                                  fill: { patternType: 'solid', fgColor: { rgb: 'FF0000' } },
                                  font: { color: { rgb: '000000' } }
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
        },

        // ✅ Merge screenshots vertically
        async mergeScreenshots({ inputFolder, outputFile }) {
          try {
            const imageFiles = fs.readdirSync(inputFolder)
              .filter(f => /\.(png|jpe?g)$/i.test(f))
              .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

            const buffers = await Promise.all(
              imageFiles.map(file => sharp(path.join(inputFolder, file)).ensureAlpha().toBuffer())
            );

            const images = await Promise.all(buffers.map(buf => sharp(buf).metadata()));
            const totalHeight = images.reduce((sum, img) => sum + img.height, 0);
            const width = Math.max(...images.map(img => img.width));

            const compositeList = [];
            let offsetY = 0;

            for (let i = 0; i < buffers.length; i++) {
              compositeList.push({ input: buffers[i], top: offsetY, left: 0 });
              offsetY += images[i].height;
            }

            const outputDir = path.dirname(outputFile);
            fs.mkdirSync(outputDir, { recursive: true });

            await sharp({
              create: {
                width,
                height: totalHeight,
                channels: 4,
                background: { r: 255, g: 255, b: 255, alpha: 1 }
              }
            })
              .composite(compositeList)
              .toFile(outputFile);

            console.log(`🧩 Merged screenshot saved: ${outputFile}`);

            // Optionally delete individual images after merge
            // imageFiles.forEach(file => fs.unlinkSync(path.join(inputFolder, file)));

            return true;
          } catch (err) {
            console.error('❌ Error merging screenshots:', err);
            return false;
          }
        },

        // ✅ Typography Excel reader
        async readTypographyData(filePath) {
          const workbook = new Excel.Workbook();
          await workbook.xlsx.readFile(filePath);
          const worksheet = workbook.getWorksheet('Typography');
          const data = [];
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            data.push({
              selector: row.getCell(1).text,
              property: row.getCell(2).text,
              expectedValue: row.getCell(3).text,
              viewport: row.getCell(4).text,
            });
          });
          return data;
        },

        // ✅ Save results to new Excel sheet
        async writeResultToExcel({ selector, property, expectedValue, actual, viewport, status }) {
          const filePath = path.resolve('cypress/results/typography_results.xlsx');
          const workbook = new Excel.Workbook();
          let worksheet;

          if (fs.existsSync(filePath)) {
            await workbook.xlsx.readFile(filePath);
            worksheet = workbook.getWorksheet('Results');
          }

          if (!worksheet) {
            worksheet = workbook.addWorksheet('Results');
            worksheet.addRow(['Selector', 'Property', 'Expected', 'Actual', 'Viewport', 'Status']);
          }

          worksheet.addRow([selector, property, expectedValue, actual, viewport, status]);

          await workbook.xlsx.writeFile(filePath);
          return true;
        },

        // ✅ Clean screenshot temp folder
        clearTempScreenshots({ folderPath }) {
          try {
            if (fs.existsSync(folderPath)) {
              const files = fs.readdirSync(folderPath);
              for (const file of files) {
                const filePath = path.join(folderPath, file);
                // Only delete files, not directories
                if (fs.statSync(filePath).isFile()) {
                  fs.unlinkSync(filePath);
                }
              }
              console.log(`🧹 Cleared temporary screenshots from: ${folderPath}`);
            }
            return true;
          } catch (err) {
            console.error('❌ Error clearing temporary screenshots:', err);
            return false;
          }
        },

      });

      return config;
    }
  }
});
