describe('Responsive Font Style Checker with Variants and Sheets', () => {
  let fontRules = {};
  let resultsByViewport = {};

  const viewports = [
    { name: '4k Screen', width: 2560, height: 1440 },
    { name: 'Normal Screen', width: 1920, height: 1080 },
    { name: 'Desktop', width: 1440, height: 900 },
    { name: 'Laptop', width: 1365, height: 800 },
    { name: 'Small Laptop', width: 1200, height: 800 },
    { name: 'Mobile', width: 990, height: 800 },
    { name: 'Small Mobile', width: 766, height: 800 },
    { name: 'Smaller Mobile', width: 431, height: 800 },
  ];

  // Map viewport names to Excel columns for expected font size
  const viewportToColumnMap = {
    'Desktop': 'Desktop',
    'Laptop': 'Laptop',
    'Tablet': 'Tablet',
    'Mobile': 'Mobile',
    'Small Mobile': 'SmallMobile',
    '4k Screen': 'Desktop',       // or your preference
    'Normal Screen': 'Desktop',
    'Smaller Mobile': 'SmallMobile',
  };

  function normalizeFontSize(value) {
    return value?.replace('px', '').trim();
  }

  before(() => {
    cy.task('readExcel', {
      filePath:
        'C:/Users/raj.khajanchi/Desktop/Cypress Font size/cypress/fixtures/Rehabiliation HubSpot Website - Style Guide.xlsx',
    }).then(data => {
      fontRules = {};
      data.slice(1).forEach(row => {
        const selectorRaw = row[0];
        if (!selectorRaw) return;

        const selector = selectorRaw.trim().toLowerCase();
        const variant = row[8]?.trim() || '';

        const rule = {
          fontFamily: row[1]?.trim(),
          Desktop: row[2]?.split('/')[0]?.trim(),
          Laptop: row[3]?.trim(),
          Tablet: row[4]?.trim(),
          Mobile: row[5]?.trim(),
          SmallMobile: row[6]?.trim(),
          fontWeight: row[7]?.toString().trim(),
          variant,
        };

        if (!fontRules[selector]) {
          fontRules[selector] = [];
        }
        fontRules[selector].push(rule);
      });

      console.log('Loaded fontRules:', fontRules);
    });
  });

  viewports.forEach(view => {
    it(`should check font styles on ${view.name}`, () => {
      cy.viewport(view.width, view.height);
      cy.visit('https://49126198.hs-sites-na2.com/');

      resultsByViewport[view.name] = [];

      const selectors = Object.keys(fontRules);

      cy.wrap(selectors).each(selector => {
        cy.document().then(doc => {
          const elements = doc.querySelectorAll(selector);
          if (elements.length === 0) {
            cy.log(`Selector "${selector}" not found on viewport ${view.name}.`);
            return;
          }

          const rulesForSelector = fontRules[selector];

          // For each variant rule for this selector
          rulesForSelector.forEach(expected => {
            // For each element found for this selector
            cy.wrap(elements).each(($el, index) => {
              cy.wrap($el).then($element => {
                const computedStyle = window.getComputedStyle($element[0]);

                const actual = {
                  fontSize: computedStyle.fontSize,
                  fontWeight: computedStyle.fontWeight,
                  fontFamily: computedStyle.fontFamily,
                };

                // Get expected font size for this viewport
                const columnName = viewportToColumnMap[view.name];
                const expectedFontSize = expected[columnName] || '';

                const isFontFamilyMatch = actual.fontFamily.includes(expected.fontFamily);
                const isFontSizeMatch = normalizeFontSize(actual.fontSize) === normalizeFontSize(expectedFontSize);
                const isFontWeightMatch = actual.fontWeight === expected.fontWeight;

                let mismatchDetails = [];
                if (!isFontFamilyMatch) mismatchDetails.push('Font Family');
                if (!isFontSizeMatch) mismatchDetails.push('Font Size');
                if (!isFontWeightMatch) mismatchDetails.push('Font Weight');

                const status = mismatchDetails.length === 0 ? 'Match' : `Mismatch: ${mismatchDetails.join(', ')}`;
                
                resultsByViewport[view.name].push({
                  Selector: selector,
                  Variant: expected.variant,
                  // Index: index + 1,
                  Text: $element.text().trim().slice(0, 50),
                  Status: status,
                  Expected_fontSize: expectedFontSize,
                  Actual_fontSize: actual.fontSize,
                  Actual_fontWeight: actual.fontWeight,
                  Actual_fontFamily: actual.fontFamily,
                  Expected_fontWeight: expected.fontWeight,
                  Expected_fontFamily: expected.fontFamily,
                });
              });
            });
          });
        });
      });
    });
  });

  after(() => {
    cy.task('writeExcelSheets', {
      data: resultsByViewport,
      filename: 'cypress/results/responsive-font-check-sheets.xlsx',
    });
  });
});
