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
        .filter(Boolean);

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

          cy.intercept('**/analytics.google.com/**', {}).as('analytics');
          cy.intercept('**/google-analytics.com/**', {}).as('ga');

          // Visit URL with timeout
          cy.visit(urlData.url, {
            failOnStatusCode: false,
            timeout: 60000  // Increased timeout in case of slow pages
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

              // To avoid duplicates, track texts already recorded for this selector + rule + viewport
              const recordedTexts = new Set();

              rulesForSelector.forEach(expected => {
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
                      Expected_fontSize: expected[view.name] || '-',
                      Actual_fontSize: 'Error',
                      Expected_lineHeight: expected.lineHeight || '-',
                      Actual_lineHeight: 'Error',
                      Expected_fontWeight: expected.fontWeight || '-',
                      Actual_fontWeight: 'Error',
                      Expected_fontFamily: expected.fontFamily || '-',
                      Actual_fontFamily: 'Error',
                      Variant: expected.variant || '-'
                    });
                    return;
                  }

                  const computedStyle = styleAnalysisResult.style;
                  const actualTextContent = styleAnalysisResult.textContent ? styleAnalysisResult.textContent.trim() : '';

                  // Skip empty text content
                  if (!actualTextContent) return;

                  // Use combination key to prevent duplicate entries for same selector, text, variant, viewport
                  const uniqueKey = `${selector}__${actualTextContent}__${expected.variant}__${view.name}`;
                  if (recordedTexts.has(uniqueKey)) {
                    return; // skip duplicate content
                  }
                  recordedTexts.add(uniqueKey);

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

                  const isFontSizeMatch = expectedFontSize === '-' ? true :
                    normalizeFontSize(actual.fontSize) === normalizeFontSize(expectedFontSize);
                  const isFontWeightMatch = expectedFontWeight === '-' ? true :
                    actual.fontWeight.toString() === expectedFontWeight.toString();
                  const isFontFamilyMatch = expectedFontFamily === '-' ? true :
                    actual.fontFamily.toString().toLowerCase().includes(expectedFontFamily.toString().toLowerCase());
                  const isLineHeightMatch = expectedLineHeight === '-' ? true :
                    actual.lineHeight.toString() === computedLineHeight;

                  const allMatch = isFontSizeMatch && isFontWeightMatch && isFontFamilyMatch && isLineHeightMatch;

                  // Highlight element with red background and border if mismatch
                  if (!allMatch) {
                    el.style.backgroundColor = 'rgba(255, 0, 0, 0.3)';
                    el.style.border = '2px solid red';
                  } else {
                    // Remove highlight if matches (helps in re-runs)
                    el.style.backgroundColor = '';
                    el.style.border = '';
                  }

                  resultsByViewport[view.name].push({
                    Selector: selector,
                    Status: allMatch ? 'Match' : 'Mismatch',
                    Text: actualTextContent.substring(0, 200),
                    Expected_fontSize: expectedFontSize,
                    Actual_fontSize: actual.fontSize,
                    Expected_lineHeight: computedLineHeight,
                    Actual_lineHeight: actual.lineHeight,
                    Expected_fontWeight: expectedFontWeight,
                    Actual_fontWeight: actual.fontWeight,
                    Expected_fontFamily: expectedFontFamily,
                    Actual_fontFamily: actual.fontFamily,
                    Variant: expected.variant || '-'
                  });
                });
              });
            });
          });
        })
        .then(() => {
          // After all viewports processed for this URL, write Excel
          return cy.task('writeExcelSheets', {
            filename: `./cypress/results/${urlData.name}-font-styles.xlsx`,
            data: resultsByViewport,
            sheetOrder: Object.keys(resultsByViewport)
          });
        });
      });
    });
  });
});
