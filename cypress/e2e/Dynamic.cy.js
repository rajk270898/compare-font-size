const screenshotTempDir = './cypress/screenshots/temp';

describe('Responsive Font Style Checker with Manual Scroll Option', () => {
  Cypress.config('defaultCommandTimeout', 3000000);
  Cypress.config('pageLoadTimeout', 3000000);

  let fontRules = {};
  let viewports = [];

  const normalizeFontSize = value => value?.toString().replace(/px/gi, '').trim() || '';

  const calculateLineHeight = (lineHeightValue, fontSize) => {
    if (!lineHeightValue || lineHeightValue === '-') return '-';
    if (lineHeightValue.includes('px')) return lineHeightValue.replace(/px/gi, '').trim() + 'px';
    const multiplier = parseFloat(lineHeightValue);
    const baseFontSize = parseFloat(fontSize.replace(/px/gi, ''));
    return isNaN(multiplier) || isNaN(baseFontSize)
      ? lineHeightValue
      : Math.round(multiplier * baseFontSize) + 'px';
  };

  const isResolution = val => {
    const num = parseInt(val?.toString().replace(/px/gi, ''));
    return !isNaN(num) && num > 0;
  };

  const createViewportConfig = resolution => {
    const width = parseInt(resolution.replace(/px/gi, ''));
    return {
      name: resolution,
      width,
      height: Math.min(Math.round(width * 0.75), 1440)
    };
  };

  const getComputedStyleForElement = rootElement => {
    let maxFontSize = -1;
    let bestStyle = null, bestText = '';

    const traverse = el => {
      if (!(el instanceof HTMLElement)) return;
      if (el.tagName.toLowerCase() === 'div' &&
        !Array.from(el.childNodes).some(n => n.nodeType === 3 && n.textContent.trim())) return;

      const style = window.getComputedStyle(el);
      const fontSize = parseFloat(style.fontSize);
      const text = Array.from(el.childNodes)
        .filter(n => n.nodeType === 3)
        .map(n => n.textContent.trim())
        .join(' ')
        .trim();

      if (text && fontSize > maxFontSize) {
        maxFontSize = fontSize;
        bestStyle = style;
        bestText = text;
      }

      Array.from(el.children).forEach(traverse);
    };

    traverse(rootElement);
    const style = bestStyle || window.getComputedStyle(rootElement);
    const text = bestText || rootElement.textContent.trim();
    return { style, textContent: text };
  };

  before(() => {
    cy.task('readExcel', { filePath: './cypress/fixtures/Style Guide.xlsx' }).then(data => {
      const headers = data[0];
      const resCols = headers.map((h, i) => isResolution(h) ? { index: i, resolution: String(h).trim() } : null).filter(Boolean);
      viewports = resCols.map(c => createViewportConfig(c.resolution));

      const colIdx = name => headers.findIndex(h => h?.toString().toLowerCase().includes(name));
      const fontFamilyIdx = colIdx('family');
      const fontWeightIdx = colIdx('weight');
      const lineHeightIdx = colIdx('height');
      const variantIdx = colIdx('variant');

      fontRules = {};
      data.slice(1).forEach(row => {
        const selector = row[0]?.trim().toLowerCase();
        if (!selector) return;
        const rule = {
          fontFamily: row[fontFamilyIdx]?.trim() || '-',
          fontWeight: row[fontWeightIdx]?.toString().trim() || '-',
          lineHeight: row[lineHeightIdx]?.toString().trim() || '-',
          variant: row[variantIdx]?.toString().trim() || '-'
        };
        resCols.forEach(c => {
          rule[c.resolution] = row[c.index]?.toString().split('/')[0]?.trim() || '-';
        });
        (fontRules[selector] ||= []).push(rule);
      });
    });
  });

  Cypress.on('uncaught:exception', () => false);

  it('checks font styles across URLs and viewports', () => {
    cy.fixture('urls.json').then(urlsData => {
      cy.wrap(urlsData.urls).each(urlData => {
        const allViewportResults = {};

        cy.wrap(viewports).each(view => {
          const results = [];
          const overlays = [];

          cy.viewport(view.width, view.height);

          [
            'analytics.google.com', 'google-analytics.com', 'doubleclick.net',
            'youtube.com', 'ytimg.com', 'linkedin.com', 'facebook.com',
            'fbcdn.net', 'googletagmanager.com', 'hotjar.com', 'clarity.ms'
          ].forEach(domain => {
            cy.intercept(`**/${domain}/**`, {}).as(domain.replace(/\./g, '-'));
          });

          cy.visit(urlData.url, { failOnStatusCode: false, timeout: 60000 });

          cy.document().then(doc => {
            const style = doc.createElement('style');
            style.innerHTML = `
              *, *::before, *::after {
                animation: none !important;
                transition: none !important;
                scroll-behavior: auto !important;
              }
            `;
            doc.head.appendChild(style);
            doc.querySelectorAll('img[loading="lazy"]').forEach(img => img.setAttribute('loading', 'eager'));
          });

          const baseName = `${urlData.name || 'page'}_${view.name}`;
          cy.task('clearTempScreenshots', { folderPath: screenshotTempDir });

          cy.window().then(win => {
            const height = win.document.documentElement.scrollHeight;
            const viewportHeight = win.innerHeight;
            const numScreens = Math.ceil(height / viewportHeight);

            const scrollAndCapture = index => {
              if (index >= numScreens) {
                // Scroll back to top
                cy.window().then(w => w.scrollTo(0, 0));
                return;
              }

              cy.window().then(w => {
                w.scrollTo(0, index * viewportHeight);
              });

              cy.wait(3000); // Give it time to render after scroll

              cy.screenshot(`temp/${baseName}_screenshot_${index}`, { capture: 'viewport' }).then(() => {
                scrollAndCapture(index + 1);
              });
            };

            scrollAndCapture(0);
          });

          const mergedOutputPath = `cypress/screenshots/mismatches/${baseName}.png`;
          cy.task('mergeScreenshots', {
            inputFolder: screenshotTempDir,
            outputFile: mergedOutputPath
          });

          cy.document().then(doc => {
            for (const [selector, rules] of Object.entries(fontRules)) {
              if (selector.includes('iframe')) continue;
              const elements = Array.from(doc.querySelectorAll(selector));
              if (elements.length === 0) {
                rules.forEach(rule => {
                  results.push({ selector, variant: rule.variant, expectedFontFamily: rule.fontFamily, expectedFontWeight: rule.fontWeight, expectedLineHeight: rule.lineHeight, expectedFontSize: rule[view.name] || '-', actualFontFamily: 'Not found', actualFontWeight: 'Not found', actualLineHeight: 'Not found', actualFontSize: 'Not found', match: false, textContent: '', Status: 'Selector not found' });
                });
              } else {
                rules.forEach(rule => {
                  elements.forEach(el => {
                    const comp = getComputedStyleForElement(el);
                    const actualFontSizeRaw = comp.style.fontSize || '-';
                    const actualFontSize = normalizeFontSize(actualFontSizeRaw);
                    const actualLineHeight = calculateLineHeight(comp.style.lineHeight || '-', actualFontSizeRaw);
                    const actualFontWeight = comp.style.fontWeight || '-';
                    const actualFontFamily = comp.style.fontFamily || '-';

                    const expected = {
                      fontSize: rule[view.name] || '-',
                      fontFamily: rule.fontFamily,
                      fontWeight: rule.fontWeight,
                      lineHeight: rule.lineHeight
                    };

                    const mismatchDetails = [];

                    if (expected.fontSize !== '-' && expected.fontSize !== actualFontSize)
                      mismatchDetails.push({ prop: 'font-size', expected: expected.fontSize, actual: actualFontSize });

                    if (expected.fontWeight !== '-' && expected.fontWeight !== actualFontWeight)
                      mismatchDetails.push({ prop: 'font-weight', expected: expected.fontWeight, actual: actualFontWeight });

                    if (expected.fontFamily !== '-' && !actualFontFamily.toLowerCase().includes(expected.fontFamily.toLowerCase()))
                      mismatchDetails.push({ prop: 'font-family', expected: expected.fontFamily, actual: actualFontFamily });

                    if (expected.lineHeight !== '-' && expected.lineHeight !== actualLineHeight)
                      mismatchDetails.push({ prop: 'line-height', expected: expected.lineHeight, actual: actualLineHeight });

                    const isMatch = mismatchDetails.length === 0;

                    if (!isMatch) {
                      const rect = el.getBoundingClientRect();
                      overlays.push({
                        top: rect.top + window.scrollY,
                        left: rect.left + window.scrollX,
                        width: rect.width,
                        height: rect.height,
                        selector,
                        variant: rule.variant,
                        mismatchDetails,
                        textContent: comp.textContent.slice(0, 100)
                      });
                    }

                    results.push({
                      selector,
                      variant: rule.variant,
                      Status: isMatch ? 'Match' : `Mismatch: ${mismatchDetails.map(d => d.prop).join(', ')}`,
                      expectedFontFamily: expected.fontFamily,
                      actualFontFamily,
                      expectedFontSize: expected.fontSize,
                      actualFontSize,
                      expectedLineHeight: expected.lineHeight,
                      actualLineHeight,
                      expectedFontWeight: expected.fontWeight,
                      actualFontWeight,
                      match: isMatch,
                      textContent: comp.textContent.slice(0, 100),
                    });
                  });
                });
              }
            }

            doc.querySelectorAll('.font-check-overlay, .font-check-label').forEach(el => el.remove());

            overlays.forEach(box => {
              const overlay = doc.createElement('div');
              Object.assign(overlay.style, {
                position: 'absolute',
                top: `${box.top}px`,
                left: `${box.left}px`,
                width: `${box.width}px`,
                height: `${box.height}px`,
                backgroundColor: 'rgba(0, 0, 0, 0.1)',
                pointerEvents: 'none',
                zIndex: 9999,
                border: '1px solid rgba(0, 0, 0, 0.5)'
              });
              overlay.classList.add('font-check-overlay');
              doc.body.appendChild(overlay);

              const mismatchDetails = box.mismatchDetails
                .filter(d => d.expected !== '-')
                .map(d => `${d.prop}: expected ${d.expected}, actual ${d.actual}`)
                .join(', ');

              const label = doc.createElement('div');
              label.textContent = `Mismatch: ${box.selector} (${box.variant}) [${mismatchDetails}]`;
              Object.assign(label.style, {
                position: 'absolute',
                top: `${box.top - 20}px`,
                left: `${box.left}px`,
                backgroundColor: 'rgba(255,255,255,0.85)',
                color: 'red',
                fontSize: '12px',
                padding: '2px 4px',
                zIndex: 10000,
                fontWeight: 'bold',
                borderRadius: '3px',
                maxWidth: '350px'
              });
              label.classList.add('font-check-label');
              doc.body.appendChild(label);
            });

            cy.then(() => {
              allViewportResults[view.name] = results;
            });
          });
        }).then(() => {
          const sheetOrder = Object.keys(allViewportResults).sort((a, b) => parseInt(b.replace(/px/gi, '')) - parseInt(a.replace(/px/gi, '')));
          cy.task('writeExcelSheets', {
            filename: `./cypress/results/${urlData.name}-font-styles.xlsx`,
            data: allViewportResults,
            sheetOrder
          });
        });
      });
    });
  });
});
