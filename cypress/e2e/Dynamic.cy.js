describe('Responsive Font Style Checker with Variants and Sheets', () => {
  // Increase test timeout to 5 minutes
  Cypress.config('defaultCommandTimeout', 300000);
  Cypress.config('pageLoadTimeout', 300000);

  let fontRules = {};
  const viewports = [
    { name: '4k Screen', width: 2560, height: 1440 },
    { name: 'Normal Screen', width: 1920, height: 1080 },
    { name: 'Desktop', width: 1440, height: 900 },
    { name: 'Laptop', width: 1366, height: 800 },
    { name: 'Small Laptop', width: 1200, height: 800 },
    { name: 'Tablet', width: 1024, height: 768 },
    { name: 'Small Tablet', width: 990, height: 800 },
    { name: 'Mobile', width: 766, height: 800 },
    { name: 'Small Mobile', width: 430, height: 800 },
  ];

  // Map viewport names to Excel columns for expected font size
  const viewportToColumnMap = {
    '4k Screen': 'Desktop',      
    'Normal Screen': 'Desktop',
    'Desktop': 'Desktop',
    'Laptop': 'Laptop',
    'Small Laptop': 'Laptop',
    'Tablet': 'Tablet',
    'Small Tablet': 'Tablet',
    'Mobile': 'Mobile',
    'Small Mobile': 'SmallMobile',
  };

  function normalizeFontSize(value) {
    return value?.replace('px', '').trim();
  }

  function calculateLineHeight(lineHeightValue, fontSize) {
    // If line height contains 'px', return it directly
    if (lineHeightValue?.includes('px')) {
      return normalizeFontSize(lineHeightValue);
    }
    // If it's a multiplier (like 1.3), multiply with font size
    const multiplier = parseFloat(lineHeightValue);
    if (!isNaN(multiplier) && fontSize) {
      const baseFontSize = parseFloat(normalizeFontSize(fontSize));
      if (!isNaN(baseFontSize)) {
        return Math.round(multiplier * baseFontSize).toString();
      }
    }
    return lineHeightValue;
  }

  before(() => {
    // First load the Excel data
    cy.task('readExcel', {
      filePath: './cypress/fixtures/Style Guide.xlsx',
    }).then(data => {
      fontRules = {};
      
      // Process font rules starting from row 1
      data.slice(1).forEach(row => {
        const selectorRaw = row[0];
        if (!selectorRaw) return;

        const selector = selectorRaw.trim().toLowerCase();

        const rule = {
          fontFamily: row[1]?.trim(),
          Desktop: row[2]?.split('/')[0]?.trim(), //1920
          Laptop: row[3]?.trim(), //1700
          SmallLaptop: row[4]?.trim(),//1200
          Tablet: row[5]?.trim(),//991
          SmallTablet: row[6]?.trim(),//767
          Mobile: row[7]?.trim(), //575
          SmallMobile: row[7]?.trim(),
          fontWeight: row[8]?.toString().trim(),
          lineHeight: row[9]?.toString().trim(),
          variant: row[10],
        };

        if (!fontRules[selector]) {
          fontRules[selector] = [];
        }
        fontRules[selector].push(rule);
      });

      console.log('Loaded fontRules:', fontRules);
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
        cy.log(`Testing URL ${urlIndex + 1}/${urls.length}: ${urlData.name}`);
        let resultsByViewport = {};

        // Test each viewport for this URL
        viewports.forEach((view, viewIndex) => {
          cy.log(`Testing viewport ${viewIndex + 1}/${viewports.length}: ${view.name}`);
          
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
            cy.log(`Processing selector ${selectorIndex + 1}/${selectors.length}: ${selector}`);
            
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
                cy.log(`Selector "${selector}" not found on viewport ${view.name}.`);
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
                        cy.log(`Element not found for selector "${selector}"`);
                        return;
                      }
                      computedStyle = window.getComputedStyle($element[0]);
                      if (!computedStyle) {
                        cy.log(`Could not get computed style for selector "${selector}"`);
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

                    const columnName = viewportToColumnMap[view.name];
                    const expectedFontSize = expected[columnName] || '-';
                    const expectedLineHeight = expected.lineHeight ? calculateLineHeight(expected.lineHeight, expectedFontSize) : '-';
                    const actualLineHeight = normalizeFontSize(actual.lineHeight);
                    const expectedFontWeight = expected.fontWeight || '-';
                    const expectedFontFamily = expected.fontFamily || '-';

                    // Only compare if expected values exist (not '-')
                    const isFontFamilyMatch = expectedFontFamily === '-' ? true : 
                      actual.fontFamily.toLowerCase().includes(expectedFontFamily.toLowerCase());
                    const isFontSizeMatch = expectedFontSize === '-' ? true : normalizeFontSize(actual.fontSize) === normalizeFontSize(expectedFontSize);
                    const isFontWeightMatch = expectedFontWeight === '-' ? true : actual.fontWeight.toString() === expectedFontWeight.toString();
                    const isLineHeightMatch = expectedLineHeight === '-' ? true : actualLineHeight === expectedLineHeight;

                    let mismatchDetails = [];
                    if (!isFontFamilyMatch) mismatchDetails.push('Font Family');
                    if (!isFontSizeMatch) mismatchDetails.push('Font Size');
                    if (!isLineHeightMatch) mismatchDetails.push('Line Height');
                    if (!isFontWeightMatch) mismatchDetails.push('Font Weight');

                    const status = mismatchDetails.length === 0 ? 'Match' : `Mismatch: ${mismatchDetails.join(', ')}`;
                    
                    resultsByViewport[view.name].push({
                      Selector: selector,
                      Variant: expected.variant || '-',
                      Text: $element.text().trim().slice(0, 50),
                      Status: status,
                      Expected_fontSize: expectedFontSize === '-' ? 'Not Found' : expectedFontSize,
                      Actual_fontSize: actual.fontSize,
                      Expected_lineHeight: expected.lineHeight ? `${expected.lineHeight} (Computed: ${expectedLineHeight}px)` : 'Not Found',
                      Actual_lineHeight: actual.lineHeight,
                      Expected_fontWeight: expectedFontWeight === '-' ? 'Not Found' : expectedFontWeight,
                      Actual_fontWeight: actual.fontWeight,
                      Expected_fontFamily: expectedFontFamily === '-' ? 'Not Found' : expectedFontFamily,
                      Actual_fontFamily: actual.fontFamily
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
