const screenshotTempDir = './cypress/screenshots/temp';

describe('Responsive Font Style Checker with Manual Scroll Option', () => {
  Cypress.config('defaultCommandTimeout', 3000000);
  Cypress.config('pageLoadTimeout', 3000000);

  let fontRules = {};
  let viewports = [];

  const LABEL_MAX_WIDTH = 450;
  const LABEL_HEIGHT = 20;

  const normalizeFontSize = value => {
    const num = parseFloat(value);
    return isNaN(num) ? '-' : `${num}px`;
  };

  const normalizeFontWeight = value => {
    if (!value || value === '-') return '-';
    const map = { normal: '400', bold: '700' };
    const lower = value.toString().trim().toLowerCase();
    return map[lower] || value;
  };

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
    let bestMatch = null;

    const traverse = el => {
      if (!(el instanceof HTMLElement)) return;

      const tag = el.tagName.toLowerCase();
      const style = window.getComputedStyle(el);
      const fontSize = parseFloat(style.fontSize);
      const text = Array.from(el.childNodes)
        .filter(n => n.nodeType === 3)
        .map(n => n.textContent.trim())
        .join(' ')
        .trim();

      if (!text) return;

      const isSemanticTag = /h[1-6]|p|button|span|label|strong/.test(tag);
      const isCurrentSemantic = bestMatch && /h[1-6]|p|button|span|label|strong/.test(bestMatch.tagName.toLowerCase());

      if (
        (fontSize > maxFontSize || (fontSize === maxFontSize && isSemanticTag)) &&
        (!bestMatch || isSemanticTag || !isCurrentSemantic)
      ) {
        maxFontSize = fontSize;
        bestMatch = el;
      }

      Array.from(el.children).forEach(traverse);
    };

    traverse(rootElement);
    const style = bestMatch ? window.getComputedStyle(bestMatch) : window.getComputedStyle(rootElement);
    const text = bestMatch ? bestMatch.textContent.trim() : rootElement.textContent.trim();
    return { style, textContent: text };
  };

  const LABEL_SPACING = 80;
  const usedAreas = new Set();
  const isPositionAvailable = (top, left, width, height) => {
    const areaKey = `${Math.round(top)},${Math.round(left)},${Math.round(width)},${Math.round(height)}`;
    if (usedAreas.has(areaKey)) return false;
    for (const area of usedAreas) {
      const [existingTop, existingLeft, existingWidth, existingHeight] = area.split(',').map(Number);
      if (!(left + width <= existingLeft ||
            existingLeft + existingWidth <= left ||
            top + height <= existingTop - LABEL_SPACING ||
            existingTop + existingHeight <= top - LABEL_SPACING)) {
        return false;
      }
    }
    return true;
  };

  const injectOverlays = (doc, overlays) => {
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
      const labelDiv = doc.createElement('div');
      let previewText = box.textContent ? box.textContent.trim().split(/\s+/).slice(0, 20).join(' ') : '';

      const pageHeight = document.documentElement.clientHeight;
      const pageWidth = document.documentElement.clientWidth;
      const labelWidth = Math.min(LABEL_MAX_WIDTH, pageWidth - 2 * LABEL_SPACING);
      const labelHeight = LABEL_HEIGHT;
      const positionsToTry = [
        { top: box.top - labelHeight - 8, left: box.left },
        { top: box.top + box.height + 8, left: box.left },
        { top: box.top, left: box.left + box.width + 8 },
        { top: box.top, left: box.left - labelWidth - 8 }
      ];

      let labelTop = 0, labelLeft = 0, found = false;
      for (const pos of positionsToTry) {
        let t = Math.max(0, Math.min(pos.top, pageHeight - labelHeight));
        let l = Math.max(LABEL_SPACING, Math.min(pos.left, pageWidth - labelWidth - LABEL_SPACING));
        if (isPositionAvailable(t, l, labelWidth, labelHeight)) {
          labelTop = t;
          labelLeft = l;
          found = true;
          break;
        }
      }
      if (!found) {
        labelTop = Math.max(0, Math.min(box.top, pageHeight - labelHeight));
        labelLeft = Math.max(LABEL_SPACING, Math.min(box.left, pageWidth - labelWidth - LABEL_SPACING));
      }
      const areaKey = `${Math.round(labelTop)},${Math.round(labelLeft)},${labelWidth},${labelHeight}`;
      usedAreas.add(areaKey);

      labelDiv.textContent = `Mismatch: ${box.selector} (${box.variant}) [${mismatchDetails}]\nPreview: ${previewText}`;
      Object.assign(labelDiv.style, {
        position: 'absolute',
        top: `${labelTop}px`,
        left: `${labelLeft}px`,
        minWidth: '120px',
        maxWidth: `${labelWidth}px`,
        width: 'fit-content',
        backgroundColor: 'rgba(255,255,255,0.95)',
        color: '#cc0000',
        fontSize: '12px',
        fontWeight: '500',
        padding: '8px 12px',
        borderRadius: '6px',
        zIndex: '10000',
        pointerEvents: 'none',
        whiteSpace: 'pre-line',
        boxShadow: '0 4px 16px rgba(0,0,0,0.18)',
        lineHeight: '1.5',
        border: '2px solid rgba(204, 0, 0, 0.25)'
      });
      overlay.appendChild(labelDiv);
    });
  };

  before(() => {
    cy.task('readExcel', { filePath: './cypress/fixtures/Style Guide.xlsx' }).then(data => {
      const headers = data[0];
      const resCols = headers
        .map((h, i) => isResolution(h) ? { index: i, resolution: String(h).trim() } : null)
        .filter(Boolean);

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

        const rawLineHeight = row[lineHeightIdx]?.toString().trim() || '-';

        const rule = {
          fontFamily: row[fontFamilyIdx]?.trim() || '-',
          fontWeight: row[fontWeightIdx]?.toString().trim() || '-',
          lineHeight: rawLineHeight,
          variant: row[variantIdx]?.toString().trim() || '-'
        };

        resCols.forEach(c => {
          const fontSize = row[c.index]?.toString().split('/')[0]?.trim() || '-';
          let expectedLineHeight = rawLineHeight;

          if (!expectedLineHeight.includes('px') && expectedLineHeight !== '-' && fontSize !== '-') {
            const num = parseFloat(expectedLineHeight);
            const size = parseFloat(fontSize);
            if (!isNaN(num) && !isNaN(size)) {
              expectedLineHeight = Math.round(num * size) + 'px';
            }
          }

          rule[c.resolution] = fontSize;
          rule[`lineHeight_${c.resolution}`] = expectedLineHeight;
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

          cy.document().then(doc => {
            for (const [selector, rules] of Object.entries(fontRules)) {
              if (selector.includes('iframe')) continue;
              const elements = Array.from(doc.querySelectorAll(selector));
              if (elements.length === 0) {
                rules.forEach(rule => {
                  results.push({
                    selector,
                    variant: rule.variant,
                    expectedFontFamily: rule.fontFamily,
                    expectedFontWeight: rule.fontWeight,
                    expectedLineHeight: rule.lineHeight,
                    expectedFontSize: rule[view.name] || '-',
                    actualFontFamily: 'Not found',
                    actualFontWeight: 'Not found',
                    actualLineHeight: 'Not found',
                    actualFontSize: 'Not found',
                    match: false,
                    textContent: '',
                    Status: 'Selector not found'
                  });
                });
              } else {
                rules.forEach(rule => {
                  elements.forEach(el => {
                    const expectedFontSize = normalizeFontSize(rule[view.name] || '-');
                    const expectedLineHeight = calculateLineHeight(rule[`lineHeight_${view.name}`], expectedFontSize);
                    const expectedFontWeight = normalizeFontWeight(rule.fontWeight || '-');
                    const expectedFontFamily = rule.fontFamily || '-';

                    const comp = getComputedStyleForElement(el);

                    const actualFontSizeRaw = comp.style.fontSize || '-';
                    const actualFontSize = normalizeFontSize(actualFontSizeRaw);
                    const actualLineHeight = calculateLineHeight(comp.style.lineHeight || '-', actualFontSizeRaw);
                    const actualFontWeight = normalizeFontWeight(comp.style.fontWeight || '-');
                    const actualFontFamily = comp.style.fontFamily || '-';
                    
                    const expected = {
                      fontSize: rule[view.name] || '-',
                      fontFamily: rule.fontFamily,
                      fontWeight: rule.fontWeight,
                      lineHeight: expectedLineHeight
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

            injectOverlays(doc, overlays);
          });

          cy.window().then(win => {
            const documentHeight = win.document.documentElement.scrollHeight;
            const viewportHeight = win.innerHeight;
            const overlap = 100;

            const positions = [];
            let scrollY = 0;

            while (scrollY < documentHeight) {
              positions.push(scrollY);
              scrollY += (viewportHeight - overlap);
            }

            const scrollAndCapture = index => {
              if (index >= positions.length) {
                cy.window().then(w => w.scrollTo(0, 0));
                return;
              }

              cy.window().then(w => {
                w.scrollTo(0, positions[index]);
              });

              cy.wait(3000);

            const paddedIndex = String(index).padStart(3, '0');
              cy.screenshot(`temp/${baseName}_screenshot_${paddedIndex}`, { capture: 'viewport' }).then(() => {
                scrollAndCapture(index + 1);
              });
            };

            scrollAndCapture(0);
          });

          const mergedOutputPath = `./cypress/screenshots/mismatches/${baseName}.png`;
          cy.task('mergeScreenshots', {
            inputFolder: screenshotTempDir,
            outputFile: mergedOutputPath
          });

          cy.then(() => {
            allViewportResults[view.name] = results;
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