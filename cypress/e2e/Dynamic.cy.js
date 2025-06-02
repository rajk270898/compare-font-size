describe('Responsive Font Style Checker with Variants and Sheets', () => {
  Cypress.config('defaultCommandTimeout', 3000000);
  Cypress.config('pageLoadTimeout', 3000000);

  let fontRules = {};
  let resolutionColumns = [];
  let viewports = [];

  function normalizeFontSize(value) {
    if (!value) return '';
    return value.toString().replace(/px/gi, '').trim();
  }

  function calculateLineHeight(lineHeightValue, fontSize) {
    if (!lineHeightValue || lineHeightValue === '-') return '-';
    if (lineHeightValue.toString().includes('px')) {
      return lineHeightValue.toString().replace(/px/gi, '').trim() + 'px';
    }
    const multiplier = parseFloat(lineHeightValue);
    if (!isNaN(multiplier) && fontSize) {
      const baseFontSize = parseFloat(fontSize.toString().replace(/px/gi, '').trim());
      if (!isNaN(baseFontSize)) {
        return Math.round(multiplier * baseFontSize).toString() + 'px';
      }
    }
    return lineHeightValue.toString();
  }

  function isResolution(value) {
    if (!value) return false;
    const cleaned = value.toString().replace(/px/gi, '').trim();
    const num = parseInt(cleaned);
    return !isNaN(num) && num > 0;
  }

  function createViewportConfig(resolution) {
    const width = parseInt(resolution.replace(/px/gi, ''));
    const height = Math.min(Math.round(width * 0.75), 1440);
    return {
      name: resolution,
      width: width,
      height: height
    };
  }

  function getComputedStyleForElement(rootElement) {
    if (!rootElement) {
      console.error('getComputedStyleForElement called with null rootElement');
      return null;
    }

    let maxFontSize = -1;
    let styleOfElementWithMaxFont = null;
    let textContentOfElementWithMaxFont = '';
    let elementWithMaxFontDetails = {};

    function findStyleRecursive(currentElement) {
      if (!(currentElement instanceof HTMLElement)) return;

      // Skip blank divs or elements without visible direct text nodes
      if (
        currentElement.tagName.toLowerCase() === 'div' &&
        !Array.from(currentElement.childNodes).some(
          (node) =>
            node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0
        )
      ) {
        return;
      }

      const style = window.getComputedStyle(currentElement);
      const fontSizeString = style.fontSize;

      let directTextContent = '';
      for (const node of currentElement.childNodes) {
        if (node.nodeType === Node.TEXT_NODE) {
          directTextContent += node.textContent.trim() + ' ';
        }
      }
      directTextContent = directTextContent.trim();

      const hasDirectVisibleText = directTextContent.length > 0;

      if (hasDirectVisibleText) {
        const fontSize = parseFloat(fontSizeString);
        if (!isNaN(fontSize) && fontSize > maxFontSize) {
          maxFontSize = fontSize;
          styleOfElementWithMaxFont = style;
          textContentOfElementWithMaxFont = directTextContent;
          elementWithMaxFontDetails = {
            tag: currentElement.tagName,
            id: currentElement.id,
            classes: currentElement.className ? currentElement.className.toString() : '',
            text: directTextContent.substring(0, 100),
            fontSize: fontSizeString
          };
        }
      }

      for (const child of currentElement.children) {
        if (child instanceof HTMLElement) {
          findStyleRecursive(child);
        }
      }
    }

    findStyleRecursive(rootElement);

    if (styleOfElementWithMaxFont) {
      return {
        style: styleOfElementWithMaxFont,
        textContent: textContentOfElementWithMaxFont,
        elementTag: elementWithMaxFontDetails.tag,
        elementClasses: elementWithMaxFontDetails.classes
      };
    } else {
      const rootStyle = window.getComputedStyle(rootElement);
      const rootText = rootElement.textContent.trim();
      return {
        style: rootStyle,
        textContent: rootText,
        elementTag: rootElement.tagName,
        elementClasses: rootElement.className ? rootElement.className.toString() : ''
      };
    }
  }

  before(() => {
    // Create screenshots directory if it doesn't exist
    cy.task('ensureDir', 'cypress/screenshots/mismatches').then(() => {
      cy.task('readExcel', {
        filePath: './cypress/fixtures/Style Guide.xlsx',
      }).then(data => {
        fontRules = {};

        const headers = data[0];

        resolutionColumns = headers
          .map((header, index) => {
            if (isResolution(header)) {
              return { index: index, resolution: header.toString().trim() };
            }
            return null;
          })
          .filter(Boolean)
          .sort((a, b) => {
            // Extract numbers from resolution strings
            const aNum = parseInt(a.resolution.replace(/px/gi, ''));
            const bNum = parseInt(b.resolution.replace(/px/gi, ''));
            // Sort in descending order (largest first)
            return bNum - aNum;
          });

        viewports = resolutionColumns.map(col => createViewportConfig(col.resolution));

        const fontFamilyIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('family'));
        const fontWeightIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('weight'));
        const lineHeightIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('height'));
        const variantIndex = headers.findIndex(h => h?.toString().toLowerCase().includes('variant'));

        data.slice(1).forEach(row => {
          const selectorRaw = row[0];
          if (!selectorRaw) return;

          const selector = selectorRaw.trim().toLowerCase();

          const rule = {
            fontFamily: row[fontFamilyIndex]?.trim() || '-',
            fontWeight: row[fontWeightIndex]?.toString().trim() || '-',
            lineHeight: row[lineHeightIndex]?.toString().trim() || '-',
            variant: row[variantIndex]?.toString().trim() || '-'
          };

          resolutionColumns.forEach(col => {
            const value = row[col.index]?.toString().split('/')[0]?.trim() || '-';
            rule[col.resolution] = value;
          });

          if (!fontRules[selector]) fontRules[selector] = [];
          fontRules[selector].push(rule);
        });

        cy.log('Loaded fontRules and viewports');
      });
    });
  });

  Cypress.on('uncaught:exception', (err) => {
    // Ignore timeout errors or other exceptions to avoid test failure
    if (err.message && err.message.includes('Timeout (b)')) {
      return false;
    }
    return false;
  });

  it('should test all URLs with all viewports and font rules', () => {
    cy.fixture('urls.json').then(urlsData => {
      const urls = urlsData.urls;

      // Process URLs sequentially to avoid race conditions
      cy.wrap(urls).each(urlData => {
        const resultsByViewport = {};

        // For each viewport
        return cy.wrap(viewports).each(view => {
          cy.viewport(view.width, view.height);

          // Block various analytics and third-party requests
          cy.intercept('**/analytics.google.com/**', {}).as('analytics');
          cy.intercept('**/google-analytics.com/**', {}).as('ga');
          cy.intercept('**/doubleclick.net/**', {}).as('doubleclick');
          cy.intercept('**/youtube.com/**', {}).as('youtube');
          cy.intercept('**/ytimg.com/**', {}).as('ytimg');
          cy.intercept('**/linkedin.com/**', {}).as('linkedin');
          cy.intercept('**/facebook.com/**', {}).as('facebook');
          cy.intercept('**/fbcdn.net/**', {}).as('fbcdn');
          cy.intercept('**/google-analytics.com/collect**', {}).as('ga-collect');
          cy.intercept('**/google-analytics.com/g/collect**', {}).as('ga4');
          cy.intercept('**/googletagmanager.com/**', {}).as('gtm');
          cy.intercept('**/hotjar.com/**', {}).as('hotjar');
          cy.intercept('**/clarity.ms/**', {}).as('clarity');

          // Visit URL with timeout
          cy.visit(urlData.url, {
            failOnStatusCode: false,
            timeout: 60000,  // Increased timeout in case of slow pages
            // Add onBeforeLoad to block more scripts
            onBeforeLoad: (win) => {
              // Block Google Analytics
              win.ga = () => {};
              win.gtag = () => {};
              // Block other tracking scripts
              win.fbq = () => {};
              win.clarity = () => {};
              win.hj = () => {};
            }
          });

          cy.wait(2000);
          cy.document().its('readyState').should('eq', 'complete');

          resultsByViewport[view.name] = [];

          const selectors = Object.keys(fontRules);

          // Process each selector sequentially
          return cy.wrap(selectors).each(selector => {
            return cy.document().then(doc => {
              const elements = doc.querySelectorAll(selector);

              // If no elements found, record Not Found once per selector per viewport
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
                return;
              }

              const rulesForSelector = fontRules[selector];

              // Process each element
              cy.wrap(elements).each(($el) => {
                const el = $el[0];
                // Skip blank divs or empty text elements
                if (
                  el.tagName.toLowerCase() === 'div' &&
                  !Array.from(el.childNodes).some(
                    (node) =>
                      node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0
                  )
                ) {
                  return; // skip blank div
                }

                // Get computed styles and text content from element with max font size inside
                const styleAnalysisResult = getComputedStyleForElement(el);

                if (!styleAnalysisResult || !styleAnalysisResult.style) {
                  resultsByViewport[view.name].push({
                    Selector: selector,
                    Status: 'Error: Style Not Found',
                    Text: $el.text().trim().substring(0, 200),
                    Expected_fontSize: '-',
                    Actual_fontSize: 'Error',
                    Expected_lineHeight: '-',
                    Actual_lineHeight: 'Error',
                    Expected_fontWeight: '-',
                    Actual_fontWeight: 'Error',
                    Expected_fontFamily: '-',
                    Actual_fontFamily: 'Error',
                    Variant: '-'
                  });
                  return;
                }

                const computedStyle = styleAnalysisResult.style;
                const actualTextContent = styleAnalysisResult.textContent ? styleAnalysisResult.textContent.trim() : '';

                // Skip empty text content
                if (!actualTextContent) return;

                // Compare against each rule variant
                rulesForSelector.forEach(expected => {
                  const actual = {
                    fontSize: computedStyle.fontSize || '',
                    lineHeight: computedStyle.lineHeight || '',
                    fontWeight: computedStyle.fontWeight || '',
                    fontFamily: computedStyle.fontFamily || ''
                  };

                  const expectedFontSize = expected[view.name] || '-';
                  const expectedLineHeight = expected.lineHeight || '-';
                  const computedLineHeight = calculateLineHeight(expectedLineHeight, expectedFontSize);
                  const expectedFontWeight = expected.fontWeight || '-';
                  const expectedFontFamily = expected.fontFamily || '-';

                  // Check each property individually
                  const mismatches = [];
                  
                  if (expectedFontSize !== '-' && 
                      normalizeFontSize(actual.fontSize) !== normalizeFontSize(expectedFontSize)) {
                    mismatches.push('Font-size');
                  }
                  
                  if (expectedLineHeight !== '-' && 
                      actual.lineHeight.toString() !== computedLineHeight) {
                    mismatches.push('Line-height');
                  }
                  
                  if (expectedFontWeight !== '-' && 
                      actual.fontWeight.toString() !== expectedFontWeight.toString()) {
                    mismatches.push('Font-weight');
                  }
                  
                  if (expectedFontFamily !== '-' && 
                      !actual.fontFamily.toString().toLowerCase().includes(expectedFontFamily.toString().toLowerCase())) {
                    mismatches.push('Font-family');
                  }

                  // Determine status message
                  let status = 'Match';
                  let screenshotPath = '';
                  
                  if (mismatches.length > 0) {
                    status = `Mismatch: ${mismatches.join(', ')}`;
                    
                    // Add visual indicators for mismatches
                    el.style.backgroundColor = 'rgba(255, 0, 0, 0.2)';
                    el.style.border = '2px solid red';
                    
                    // Ensure element is in view with padding
                    el.scrollIntoView({ behavior: 'instant', block: 'center' });
                    cy.wait(500); // Wait for scrolling to complete

                    // Create text overlay using Cypress command
                    cy.window().then((win) => {
                      // Create detailed mismatch text
                      const mismatchDetails = mismatches.map(type => {
                        let details = '';
                        switch(type) {
                          case 'Font-size':
                            details = `Expected: ${expectedFontSize} | Found: ${actual.fontSize}`;
                            break;
                          case 'Line-height':
                            details = `Expected: ${computedLineHeight} | Found: ${actual.lineHeight}`;
                            break;
                          case 'Font-weight':
                            details = `Expected: ${expectedFontWeight} | Found: ${actual.fontWeight}`;
                            break;
                          case 'Font-family':
                            details = `Expected: ${expectedFontFamily} | Found: ${actual.fontFamily}`;
                            break;
                        }
                        return `${type}: ${details}`;
                      }).join('\\n');

                      const overlayText = `Selector: ${selector}\\nVariant: ${expected.variant || 'default'}\\n${mismatchDetails}`;
                      
                      // Create overlay element
                      const overlay = win.document.createElement('div');
                      overlay.id = 'mismatch-overlay';
                      overlay.style.position = 'fixed';
                      overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.8)';
                      overlay.style.color = '#FF4444';
                      overlay.style.padding = '10px 15px';
                      overlay.style.fontSize = '16px';
                      overlay.style.fontWeight = 'bold';
                      overlay.style.borderRadius = '5px';
                      overlay.style.zIndex = '999999';
                      overlay.style.maxWidth = '600px';
                      overlay.style.whiteSpace = 'pre-wrap';
                      overlay.textContent = overlayText;

                      // Get element position
                      const rect = el.getBoundingClientRect();
                      overlay.style.top = (rect.top + win.scrollY - 100) + 'px';
                      overlay.style.left = rect.left + 'px';

                      // Add overlay to page
                      win.document.body.appendChild(overlay);

                      // Generate screenshot name and path
                      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
                      const screenshotName = `${selector}_${expected.variant || 'default'}_${timestamp}`;
                      const urlFolderName = urlData.name.replace(/[^a-zA-Z0-9-_]/g, '_');
                      const screenshotRelativePath = `mismatches/${urlFolderName}/${view.name}`;
                      const finalScreenshotPath = `cypress/screenshots/${screenshotRelativePath}/${screenshotName}.png`;

                      // Ensure URL-specific directory exists
                      cy.task('ensureDir', `cypress/screenshots/${screenshotRelativePath}`).then(() => {
                        // Delete existing screenshots in this directory
                        cy.task('deleteFolder', `cypress/screenshots/${screenshotRelativePath}`).then(() => {
                          cy.screenshot(`${screenshotRelativePath}/${screenshotName}`, {
                            capture: 'viewport',
                            clip: {
                              x: Math.max(0, rect.left - 700),
                              y: Math.max(0, rect.top - 700),
                              width: rect.width + 1400,
                              height: rect.height + 1400
                            },
                            scale: false,
                            overwrite: true
                          }).then(() => {
                            // Remove overlay after screenshot
                            win.document.body.removeChild(overlay);
                          });
                        });
                      });

                      // Add result with screenshot path
                      resultsByViewport[view.name].push({
                        Selector: selector,
                        Status: status,
                        Text: actualTextContent.substring(0, 200),
                        Expected_fontSize: expectedFontSize,
                        Actual_fontSize: actual.fontSize,
                        Expected_lineHeight: computedLineHeight,
                        Actual_lineHeight: actual.lineHeight,
                        Expected_fontWeight: expectedFontWeight,
                        Actual_fontWeight: actual.fontWeight,
                        Expected_fontFamily: expectedFontFamily,
                        Actual_fontFamily: actual.fontFamily,
                        Variant: expected.variant || '-',
                        Screenshot: finalScreenshotPath
                      });
                    });
                  } else {
                    // Add result without screenshot for matches
                    resultsByViewport[view.name].push({
                      Selector: selector,
                      Status: status,
                      Text: actualTextContent.substring(0, 200),
                      Expected_fontSize: expectedFontSize,
                      Actual_fontSize: actual.fontSize,
                      Expected_lineHeight: computedLineHeight,
                      Actual_lineHeight: actual.lineHeight,
                      Expected_fontWeight: expectedFontWeight,
                      Actual_fontWeight: actual.fontWeight,
                      Expected_fontFamily: expectedFontFamily,
                      Actual_fontFamily: actual.fontFamily,
                      Variant: expected.variant || '-',
                      Screenshot: ''
                    });
                  }
                });
              });
            });
          });
        })
        .then(() => {
          // After all viewports processed for this URL, write Excel
          // Sort viewports to ensure 1920px comes first
          const sortedViewportNames = Object.keys(resultsByViewport).sort((a, b) => {
            const aNum = parseInt(a.replace(/px/gi, ''));
            const bNum = parseInt(b.replace(/px/gi, ''));
            return bNum - aNum;  // Sort in descending order
          });

          return cy.task('writeExcelSheets', {
            filename: `./cypress/results/${urlData.name}-font-styles.xlsx`,
            data: resultsByViewport,
            sheetOrder: sortedViewportNames
          });
        });
      });
    });
  });
});






