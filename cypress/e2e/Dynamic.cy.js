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

    // Constants for label layout
    const LABEL_HEIGHT = 24;
    const LABEL_SPACING = 8;
    const LABEL_MAX_WIDTH = 450;
    const VIEWPORT_PADDING = 10;

    const pageWidth = doc.documentElement.clientWidth;
    const pageHeight = doc.documentElement.clientHeight;
    const usedAreas = new Set();

    const isPositionAvailable = (top, left, width, height) => {
      const box = `${Math.round(top)},${Math.round(left)},${Math.round(width)},${Math.round(height)}`;
      if (usedAreas.has(box)) return false;

      for (const area of usedAreas) {
        const [existingTop, existingLeft, existingWidth, existingHeight] = area.split(',').map(Number);
        if (!(left + width <= existingLeft - LABEL_SPACING ||
              existingLeft + existingWidth <= left - LABEL_SPACING ||
              top + height <= existingTop - LABEL_SPACING ||
              existingTop + existingHeight <= top - LABEL_SPACING)) {
          return false;
        }
      }
      return true;
    };

    // Group overlays by their element position to handle multiple variants
    const groupedOverlays = {};
    overlays.forEach(box => {
      const key = `${Math.round(box.top)},${Math.round(box.left)},${Math.round(box.width)},${Math.round(box.height)}`;
      if (!groupedOverlays[key]) {
        groupedOverlays[key] = [];
      }
      groupedOverlays[key].push(box);
    });

    Object.values(groupedOverlays).forEach(boxGroup => {
      const firstBox = boxGroup[0];
      const overlay = doc.createElement('div');
      Object.assign(overlay.style, {
        position: 'absolute',
        top: `${firstBox.top}px`,
        left: `${firstBox.left}px`,
        width: `${firstBox.width}px`,
        height: `${firstBox.height}px`,
        backgroundColor: 'rgba(0, 0, 0, 0.6)',
        pointerEvents: 'none',
        zIndex: 9998
      });
      overlay.classList.add('font-check-overlay');
      doc.body.appendChild(overlay);

      // Create labels for each variant
      boxGroup.forEach((box, index) => {
        const mismatchDetails = box.mismatchDetails
          .filter(d => d.expected !== '-')
          .map(d => {
            // Shorten property names but keep expected/actual format
            const prop = d.prop.replace('font-', '').replace('line-', '');
            return `${prop}: expected ${d.expected}, actual ${d.actual}`;
          })
          .join(' | ');

        const labelDiv = doc.createElement('div');
        
        // Calculate position for stacked labels
        let labelTop = box.top - LABEL_HEIGHT - 2 - (index * (LABEL_HEIGHT + 4));
        let labelLeft = box.left;

        // If labels would be off-screen at the top, position them below the element
        if (labelTop < VIEWPORT_PADDING) {
          labelTop = box.top + box.height + 2 + (index * (LABEL_HEIGHT + 4));
        }

        // Adjust horizontal position to keep label within viewport
        if (labelLeft + LABEL_MAX_WIDTH > pageWidth - VIEWPORT_PADDING) {
          labelLeft = Math.max(VIEWPORT_PADDING, pageWidth - LABEL_MAX_WIDTH - VIEWPORT_PADDING);
        }

        Object.assign(labelDiv.style, {
          position: 'absolute',
          top: `${labelTop}px`,
          left: `${labelLeft}px`,
          width: 'auto',
          maxWidth: `${LABEL_MAX_WIDTH}px`,
          height: `${LABEL_HEIGHT}px`,
          backgroundColor: 'rgba(255, 255, 255, 0.98)',
          color: '#cc0000',
          fontSize: '11px',
          fontWeight: '500',
          padding: '4px 8px',
          borderRadius: '3px',
          zIndex: 9999,
          pointerEvents: 'none',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          lineHeight: '16px',
          display: 'flex',
          alignItems: 'center',
          transform: 'translateY(0)',
          opacity: '1'
        });

        // Include variant information in the label if it exists
        const variantInfo = box.variant ? ` (${box.variant})` : '';
        labelDiv.textContent = `Mismatch: ${box.selector}${variantInfo} | ${mismatchDetails}`;
        labelDiv.classList.add('font-check-label');
        doc.body.appendChild(labelDiv);

        // Register the used area
        const actualWidth = Math.min(labelDiv.offsetWidth + 16, LABEL_MAX_WIDTH); // Add padding to width
        const areaKey = `${Math.round(labelTop)},${Math.round(labelLeft)},${actualWidth},${LABEL_HEIGHT}`;
        usedAreas.add(areaKey);
      });
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

          cy.visit(urlData.url, { 
            failOnStatusCode: false, 
            timeout: 60000 
          });

          // Wait for page load and resources
          cy.document().then(doc => {
            return new Cypress.Promise((resolve) => {
              if (doc.readyState === 'complete') {
                resolve();
              } else {
                const onLoad = () => {
                  doc.removeEventListener('load', onLoad);
                  resolve();
                };
                doc.addEventListener('load', onLoad);
              }
            });
          });

          // Wait for all stylesheets to be loaded
          cy.document().then(doc => {
            const styleSheets = Array.from(doc.styleSheets);
            const loadingSheets = styleSheets.filter(sheet => {
              try {
                return !sheet.cssRules; // Will throw if stylesheet not loaded
              } catch (e) {
                return true;
              }
            });
            if (loadingSheets.length > 0) {
              cy.wait(2000); // Give additional time for stylesheets to load
            }
          });

          // Disable animations and transitions
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

            // Set all lazy-loaded images to eager loading
            doc.querySelectorAll('img[loading="lazy"]').forEach(img => {
              img.setAttribute('loading', 'eager');
            });

            // Wait for all images to load
            const images = Array.from(doc.getElementsByTagName('img'));
            const loadingImages = images.filter(img => !img.complete);
            
            if (loadingImages.length > 0) {
              cy.wrap(loadingImages).each($img => {
                cy.wrap($img).should('have.prop', 'complete', true);
              });
            }
          });

          // Clear any existing screenshots
          cy.task('clearTempScreenshots', { folderPath: screenshotTempDir });

          // Continue with font checking...
          const baseName = `${urlData.name || 'page'}_${view.name}`;

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
                      textContent: comp.textContent.slice(0, 100)
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
            const overlap = Math.round(viewportHeight * 0.2); // 20% overlap for smooth transitions

            const positions = [];
            let scrollY = 0;
            while (scrollY < documentHeight) {
              positions.push(scrollY);
              scrollY += (viewportHeight - overlap);
            }

            // Ensure we capture the bottom of the page
            if (positions[positions.length - 1] < documentHeight - viewportHeight) {
              positions.push(documentHeight - viewportHeight);
            }

            const performScrollAndCapture = async (positions, overlays, baseName) => {
              let currentIndex = 0;

              const scrollAndCapture = async () => {
                if (currentIndex >= positions.length) {
                  cy.window().then(w => w.scrollTo(0, 0));
                  return;
                }

                const currentScroll = positions[currentIndex];

                // Wait for any animations to complete
                cy.wait(1000);

                // Scroll to position
                cy.window().then(w => w.scrollTo(0, currentScroll));

                // Wait for scroll to settle and overlays to update
                cy.wait(1000);

                // Update overlays with current scroll position
                cy.document().then(doc => {
                  doc.querySelectorAll('.font-check-overlay, .font-check-label').forEach(el => el.remove());
                  injectOverlays(doc, overlays, currentScroll);
                });

                // Wait for overlays to be fully rendered
                cy.wait(500);

                // Take screenshot
                cy.screenshot(`temp/${baseName}_screenshot_${String(currentIndex).padStart(3, '0')}`, {
                  capture: 'viewport',
                  overwrite: true
                });

                currentIndex++;
                scrollAndCapture();
              };

              await scrollAndCapture();
            };

            performScrollAndCapture(positions, overlays, baseName).then(() => {
              // After all screenshots are taken, merge them
              cy.task('mergeScreenshots', {
                inputFolder: screenshotTempDir,
                outputFile: `./cypress/screenshots/mismatches/${baseName}.png`
              }).then(() => {
                // Clean up temporary screenshots after successful merge
                cy.task('clearTempScreenshots', { 
                  folderPath: screenshotTempDir 
                });
              });
            });
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