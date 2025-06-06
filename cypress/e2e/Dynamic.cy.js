const screenshotTempDir = './cypress/screenshots/temp';

// Full test suite for font validation
describe('Responsive Font Style Checker with Manual Scroll Option', () => {
  Cypress.config('defaultCommandTimeout', 3000000);
  Cypress.config('pageLoadTimeout', 3000000);

  let fontRules = {};
  let viewports = [];

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

function injectOverlays(doc, overlays, scrollY = 0) {
  const existingOverlay = doc.querySelector('.font-check-overlay');
  if (existingOverlay) existingOverlay.remove();

  const overlayContainer = doc.createElement('div');
  overlayContainer.className = 'font-check-overlay';
  Object.assign(overlayContainer.style, {
    position: 'absolute',
    top: '0',
    left: '0',
    width: '100%',
    height: '100%',
    pointerEvents: 'none',
    zIndex: '9999'
  });

  // Constants for layout
  const HEADER_HEIGHT = 80;
  const FOOTER_HEIGHT = 80;
  const LABEL_HEIGHT = 32;
  const LABEL_SPACING = 8;
  const LABEL_MAX_WIDTH = 450;
  const COLUMN_SPACING = 20;
  const GRID_CELL_HEIGHT = LABEL_HEIGHT + LABEL_SPACING;
  const MIN_VERTICAL_GAP = 10; // Minimum gap between labels

  // Initialize layout tracking
  const pageHeight = doc.documentElement.scrollHeight;
  const pageWidth = doc.documentElement.clientWidth;
  const usedAreas = new Set(); // Track used areas to prevent overlaps

  // Helper function to check if a position is available
  const isPositionAvailable = (top, left, width, height) => {
    const box = `${Math.round(top)},${Math.round(left)},${Math.round(width)},${Math.round(height)}`;
    if (usedAreas.has(box)) return false;
    
    // Check for overlaps with existing labels
    for (const area of usedAreas) {
      const [existingTop, existingLeft, existingWidth, existingHeight] = area.split(',').map(Number);
      if (!(left + width <= existingLeft || 
            existingLeft + existingWidth <= left || 
            top + height <= existingTop - MIN_VERTICAL_GAP || 
            existingTop + existingHeight <= top - MIN_VERTICAL_GAP)) {
        return false;
      }
    }
    return true;
  };

  // Group overlays by their vertical position
  const groupedOverlays = overlays.reduce((groups, overlay) => {
    const el = doc.querySelector(overlay.selector);
    if (!el) return groups;
    
    const rect = el.getBoundingClientRect();
    const top = rect.top + window.scrollY;
    const groupKey = Math.floor(top / 100) * 100;
    
    if (!groups[groupKey]) groups[groupKey] = [];
    groups[groupKey].push({ ...overlay, rect, el, top });
    return groups;
  }, {});

  // Process overlays in vertical order
  Object.entries(groupedOverlays)
    .sort(([a], [b]) => Number(a) - Number(b))
    .forEach(([_, sectionOverlays]) => {
      // Sort elements within each section by their DOM order
      sectionOverlays
        .sort((a, b) => a.top - b.top)
        .forEach(({ selector, mismatchDetails, rect, el }) => {
          if (!el || !rect) return;

          const mismatchText = mismatchDetails.map(
            m => `${m.prop}: expected ${m.expected}, actual ${m.actual}`
          ).join(', ');

          const label = `Mismatch: ${selector} (-) [${mismatchText}]`;
          const labelDiv = doc.createElement('div');
          labelDiv.className = 'font-check-label';
          labelDiv.textContent = label;

          // Calculate initial position
          const elementTop = rect.top + window.scrollY;
          const elementBottom = rect.bottom + window.scrollY;
          let labelTop = elementBottom + 4;
          let labelLeft = Math.min(Math.max(rect.left, LABEL_SPACING), pageWidth - LABEL_MAX_WIDTH - LABEL_SPACING);
          
          // Try different positions until we find an available spot
          let attempts = 0;
          const maxAttempts = 5;
          
          while (attempts < maxAttempts) {
            if (isPositionAvailable(labelTop, labelLeft, LABEL_MAX_WIDTH, LABEL_HEIGHT)) {
              break;
            }
            
            // Try different positions in this order:
            // 1. Move down
            // 2. Move to the right
            // 3. Try above the element
            if (attempts === 0) {
              labelTop += GRID_CELL_HEIGHT;
            } else if (attempts === 1) {
              labelLeft += LABEL_MAX_WIDTH + COLUMN_SPACING;
              labelTop = elementBottom + 4;
            } else if (attempts === 2) {
              labelTop = elementTop - LABEL_HEIGHT - 4;
              labelLeft = Math.min(Math.max(rect.left, LABEL_SPACING), pageWidth - LABEL_MAX_WIDTH - LABEL_SPACING);
            } else {
              labelTop += GRID_CELL_HEIGHT;
            }
            
            attempts++;
          }

          // If we found a valid position, place the label
          if (attempts < maxAttempts) {
            const box = `${Math.round(labelTop)},${Math.round(labelLeft)},${LABEL_MAX_WIDTH},${LABEL_HEIGHT}`;
            usedAreas.add(box);

            Object.assign(labelDiv.style, {
              position: 'absolute',
              top: `${labelTop}px`,
              left: `${labelLeft}px`,
              backgroundColor: 'rgba(248, 218, 218, 0.9)',
              color: '#cc0000',
              fontSize: '12px',
              fontWeight: '500',
              padding: '6px 10px',
              borderRadius: '4px',
              zIndex: '10000',
              pointerEvents: 'none',
              whiteSpace: 'pre-wrap',
              maxWidth: `${LABEL_MAX_WIDTH}px`,
              boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
              lineHeight: '1.4',
              border: '1px solid rgba(204, 0, 0, 0.2)'
            });

            overlayContainer.appendChild(labelDiv);
          }
        });
    });

  doc.body.appendChild(overlayContainer);
}

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
          });

          const performScrollAndCapture = (positions, overlays, baseName) => {
          const scrollAndCapture = index => {
            if (index >= positions.length) {
              cy.window().then(w => w.scrollTo(0, 0));
              return;
            }

            const currentScroll = positions[index];

            cy.window().then(w => {
              w.scrollTo(0, currentScroll);
            });

            cy.wait(1000).then(() => {
              cy.document().then(doc => {
                console.log('Injecting overlays at scrollY:', currentScroll, 'Index:', index);
                injectOverlays(doc, overlays, currentScroll);
              });
            });

            cy.wait(1500).then(() => {
              const paddedIndex = String(index).padStart(3, '0');
              cy.screenshot(`temp/${baseName}_screenshot_${paddedIndex}`, { capture: 'viewport' }).then(() => {
                cy.document().then(doc => {
                  const old = doc.querySelector('.font-check-overlay');
                  if (old) old.remove();
                });
                scrollAndCapture(index + 1);
              });
            });
          };
          scrollAndCapture(0);
        };

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

          performScrollAndCapture(positions, overlays, baseName);
        });

          const mergedOutputPath = `./cypress/screenshots/mismatches/${baseName}.png`;
          cy.task('mergeScreenshots', {
            inputFolder: screenshotTempDir,
            outputFile: mergedOutputPath
          });

          cy.then(() => {
            allViewportResults[view.name] = results;
          });
        }); // closes cy.wrap(viewports).each

      // After all viewports for this URL
      cy.then(() => {
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