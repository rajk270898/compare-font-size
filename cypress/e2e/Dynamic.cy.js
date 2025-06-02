describe('Responsive Font Style Checker with Variants and Sheets', () => {
  // Increase test timeout to 5 minutes
  Cypress.config('defaultCommandTimeout', 300000);
  Cypress.config('pageLoadTimeout', 300000);

  let fontRules = {};
  let resolutionColumns = [];
  let viewports = [];

  // Function to detect if a value represents a resolution
  function isResolution(value) {
    // Remove 'px' and try to parse as number
    const num = parseInt(value?.toString().replace(/px/gi, ''));
    return !isNaN(num) && num > 0;
  }

  function normalizeFontSize(value) {
    if (!value) return '';
    return value.toString().replace(/px/gi, '').trim();
  }

  function calculateLineHeight(lineHeightValue, fontSize) {
    if (!lineHeightValue || lineHeightValue === '-') return '-';
    
    // If line height contains 'px', return it directly
    if (lineHeightValue.toString().includes('px')) {
      return lineHeightValue.toString().replace(/px/gi, '').trim() + 'px';
    }
    
    // If it's a multiplier (like 1.3), multiply with font size
    const multiplier = parseFloat(lineHeightValue);
    if (!isNaN(multiplier) && fontSize) {
      const baseFontSize = parseFloat(fontSize.toString().replace(/px/gi, '').trim());
      if (!isNaN(baseFontSize)) {
        return Math.round(multiplier * baseFontSize).toString() + 'px';
      }
    }
    return lineHeightValue.toString();
  }

  function compareFontSizes(expected, actual) {
    const normalizedExpected = normalizeFontSize(expected);
    const normalizedActual = normalizeFontSize(actual);
    return normalizedExpected === normalizedActual;
  }

  // Function to create viewport config from resolution
  function createViewportConfig(resolution) {
    const width = parseInt(resolution.replace(/px/gi, ''));
    // Default height calculation, adjust if needed
    const height = Math.min(Math.round(width * 0.75), 1440);
    return {
      name: resolution,  // Keep the original resolution string (e.g., "1920px")
      width: width,
      height: height
    };
  }

  before(() => {
    // First load the Excel data
    cy.task('readExcel', {
      filePath: './cypress/fixtures/Style Guide.xlsx',
    }).then(data => {
      fontRules = {};
      
      // Get headers from first row and store original order
      const headers = data[0];
      const resolutionOrder = [];
      
      // Find column indices and track resolution order
      headers.forEach((header, index) => {
        if (isResolution(header)) {
          resolutionOrder.push({
            resolution: header.toString().trim(),
            index: index
          });
        }
      });
      
      // Find other column indices
      const fontWeightIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('weight'));
      const lineHeightIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('height'));
      const fontFamilyIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('family'));
      const variantIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('variant'));
      
      // Create viewports maintaining original order
      viewports = resolutionOrder.map(({ resolution }) => createViewportConfig(resolution));
      resolutionColumns = resolutionOrder;

      // Process font rules starting from row 1
      data.slice(1).forEach(row => {
        const selectorRaw = row[0];
        if (!selectorRaw) return;

        const selector = selectorRaw.trim().toLowerCase();
        const lineHeight = row[lineHeightIndex]?.toString().trim() || '-';
        
        // Create dynamic rule object with correct indices
        const rule = {
          fontFamily: row[fontFamilyIndex]?.trim() || '-',
          fontWeight: row[fontWeightIndex]?.toString().trim() || '-',
          lineHeight: lineHeight,
          variant: row[variantIndex]?.toString().trim() || '-'
        };

        // Add resolutions in original order
        resolutionOrder.forEach(({ resolution, index }) => {
          const value = row[index]?.toString().split('/')[0]?.trim() || '-';
          rule[resolution] = value;
        });

        if (!fontRules[selector]) {
          fontRules[selector] = [];
        }
        fontRules[selector].push(rule);
      });

      console.log('Loaded fontRules:', fontRules);
      console.log('Detected viewports in order:', viewports);
    });
  });

  // Add this at the top level of the describe block
  Cypress.on('uncaught:exception', (err, runnable) => {
    // Ignore Google Analytics errors
    if (err.message && err.message.includes('Timeout (b)')) {
      return false;
    }
    // returning false here prevents Cypress from failing the test
    return false;
  });

  // Load URLs and create test suites
  it('should test all URLs', () => {
    cy.fixture('urls.json').then(urlsData => {
      const urls = urlsData.urls;
      cy.log('Starting tests for URLs:', urls);

      // Process each URL
      urls.forEach((urlData, urlIndex) => {
        // cy.log(`Testing URL ${urlIndex + 1}/${urls.length}: ${urlData.name}`);
        let resultsByViewport = {};

        // Test each viewport for this URL
        viewports.forEach((view, viewIndex) => {
          // cy.log(`Testing viewport ${viewIndex + 1}/${viewports.length}: ${view.name}`);
          
          cy.viewport(view.width, view.height);
          
          // Visit the current URL and handle analytics
          cy.intercept('**/analytics.google.com/**', {}).as('analytics');
          cy.intercept('**/google-analytics.com/**', {}).as('ga');
          
          cy.visit(urlData.url, {
            failOnStatusCode: false,
            timeout: 30000
          });

          // Wait for page load and ignore analytics
          cy.wait(2000); // Give page time to start loading
          cy.document().its('readyState').should('eq', 'complete');
          
          resultsByViewport[view.name] = [];
          const selectors = Object.keys(fontRules);

          // Process selectors with progress logging
          cy.wrap(selectors).each((selector, selectorIndex) => {
            // cy.log(`Processing selector ${selectorIndex + 1}/${selectors.length}: ${selector}`);
            
            cy.document().then(doc => {
              const elements = doc.querySelectorAll(selector);
              if (elements.length === 0) {
                resultsByViewport[view.name].push({
                  Selector: selector,
                  Status: 'Not Found',
                  Text: '',
                  Expected_fontSize: '',
                  Actual_fontSize: '',
                  Expected_lineHeight: '',
                  Actual_lineHeight: '',
                  Expected_fontWeight: '',
                  Actual_fontWeight: '',
                  Expected_fontFamily: '',
                  Actual_fontFamily: '',
                  Variant: ''
                });
                // cy.log(`Selector "${selector}" not found on viewport ${view.name}.`);
                return;
              }

              const rulesForSelector = fontRules[selector];

              rulesForSelector.forEach(expected => {
                cy.wrap(elements).each(($el, index) => {
                  // Ensure element is still in DOM and wait for it to be visible
                  cy.wrap($el).should('exist').then($element => {
                    let computedStyle;
                    try {
                      if (!$element || !$element[0]) {
                        // cy.log(`Element not found for selector "${selector}"`);
                        return;
                      }
                      computedStyle = window.getComputedStyle($element[0]);
                      if (!computedStyle) {
                        // cy.log(`Could not get computed style for selector "${selector}"`);
                        return;
                      }
                    } catch (e) {
                      cy.log(`Error getting computed style for selector "${selector}":`, e);
                      return;
                    }

                    const actual = {
                      fontSize: computedStyle.fontSize || '',
                      lineHeight: computedStyle.lineHeight || '',
                      fontWeight: computedStyle.fontWeight || '',
                      fontFamily: computedStyle.fontFamily || '',
                    };

                    const columnName = viewports.find(v => v.name === view.name)?.name || '-';
                    const expectedFontSize = expected[columnName] || '-';
                    const expectedLineHeight = expected.lineHeight || '-';
                    const computedLineHeight = calculateLineHeight(expectedLineHeight, expectedFontSize);
                    const actualLineHeight = normalizeFontSize(actual.lineHeight);
                    const expectedFontWeight = expected.fontWeight || '-';
                    const expectedFontFamily = expected.fontFamily || '-';

                    // Only compare if expected values exist (not '-')
                    const isFontFamilyMatch = expectedFontFamily === '-' ? true : 
                      actual.fontFamily.toLowerCase().includes(expectedFontFamily.toLowerCase());
                    const isFontSizeMatch = expectedFontSize === '-' ? true : compareFontSizes(expectedFontSize, actual.fontSize);
                    const isFontWeightMatch = expectedFontWeight === '-' ? true : actual.fontWeight.toString() === expectedFontWeight.toString();
                    const isLineHeightMatch = expectedLineHeight === '-' ? true : 
                      normalizeFontSize(actual.lineHeight) === normalizeFontSize(computedLineHeight);

                    let mismatchDetails = [];
                    if (!isFontFamilyMatch) mismatchDetails.push('Font Family');
                    if (!isFontSizeMatch) mismatchDetails.push('Font Size');
                    if (!isLineHeightMatch) mismatchDetails.push('Line Height');
                    if (!isFontWeightMatch) mismatchDetails.push('Font Weight');

                    const status = mismatchDetails.length === 0 ? 'Match' : `Mismatch: ${mismatchDetails.join(', ')}`;
                    
                    resultsByViewport[view.name].push({
                      Selector: selector,
                      Status: status,
                      Text: $element.text().trim().slice(0, 50),
                      Expected_fontSize: expectedFontSize === '-' ? 'Not Found' : expectedFontSize,
                      Actual_fontSize: actual.fontSize,
                      Expected_lineHeight: expectedLineHeight === '-' ? 'Not Found' : computedLineHeight,
                      Actual_lineHeight: actual.lineHeight,
                      Expected_fontWeight: expectedFontWeight === '-' ? 'Not Found' : expectedFontWeight,
                      Actual_fontWeight: actual.fontWeight,
                      Expected_fontFamily: expectedFontFamily === '-' ? 'Not Found' : expectedFontFamily,
                      Actual_fontFamily: actual.fontFamily,
                      Variant: expected.variant || '-'
                    });
                  });
                });
              });
            });
          }).then(() => {
            // Log completion after all selectors are processed for this viewport
            cy.log(`Completed viewport ${view.name} for ${urlData.name}`);
          });
        });

        // Write results for this URL with clear logging
        cy.then(() => {
          const hasData = Object.values(resultsByViewport).some(data => 
            Array.isArray(data) && data.length > 0
          );

          if (hasData) {
            const safeFileName = urlData.name.toLowerCase().replace(/[^a-z0-9]/g, '-');
            const outputPath = `./cypress/results/${safeFileName}-font-check.xlsx`;
            
            cy.log(`Writing results for ${urlData.name} to ${outputPath}`);
            
            cy.task('writeExcelSheets', {
              data: resultsByViewport,
              filename: outputPath,
              sheetOrder: viewports.map(v => v.name) // Pass the correct sheet order
            }).then((result) => {
              if (result === null) {
                cy.log(`✅ Successfully wrote results for ${urlData.name}`);
              }
            }, (error) => {
              cy.log(`❌ Error writing results for ${urlData.name}: ${error.message}`);
            });
          } else {
            cy.log(`⚠️ No data to write for ${urlData.name}`);
          }
        });
      });

      // After all URLs are processed
      cy.then(() => {
        cy.log('✅ Test completed for all URLs');
      });
    });
  });
});
