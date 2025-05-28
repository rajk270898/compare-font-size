describe('Responsive Font Style Checker with Variants and Sheets', () => {
  let fontRules = {};
  let resultsByViewport = {};
  let websiteUrl = '';
  let isStyleGuideAvailable = true;
  
  const defaultSelectors = [
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'button', '.subtitle'
  ];

  const viewports = [
    { name: '4k Screen', width: 2560, height: 1440 },
    // { name: 'Normal Screen', width: 1920, height: 1080 },
    // { name: 'Desktop', width: 1440, height: 900 },
    // { name: 'Laptop', width: 1365, height: 800 },
    // { name: 'Small Laptop', width: 1200, height: 800 },
    // { name: 'Mobile', width: 990, height: 800 },
    // { name: 'Small Mobile', width: 766, height: 800 },
    // { name: 'Smaller Mobile', width: 431, height: 800 },
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
    // Read the homepage URL from urls.json
    cy.fixture('urls.json').then(urls => {
      websiteUrl = urls.homepage;
    });

    // Try to read the style guide Excel file
    cy.task('readExcel', {
      filePath: './cypress/fixtures/Rehabiliation HubSpot Website - Style Guide.xlsx',
    }).then(data => {
      if (!data) {
        isStyleGuideAvailable = false;
        return;
      }

      fontRules = {};
      
      data.slice(1).forEach(row => {
        const selectorRaw = row[0];
        if (!selectorRaw) return;

        const selector = selectorRaw.trim().toLowerCase();
        const variant = row[9]?.trim() || '';

        const rule = {
          fontFamily: row[1]?.trim(),
          Desktop: row[2]?.split('/')[0]?.trim(),
          Laptop: row[3]?.trim(),
          Tablet: row[4]?.trim(),
          Mobile: row[5]?.trim(),
          SmallMobile: row[6]?.trim(),
          fontWeight: row[7]?.toString().trim(),
          lineHeight: row[8]?.toString().trim(),
          variant,
        };

        if (!fontRules[selector]) {
          fontRules[selector] = [];
        }
        fontRules[selector].push(rule);
      });
    }, (error) => {
      // Handle error in the task
      cy.log('Style guide file not found or error reading it:', error);
      isStyleGuideAvailable = false;
    });
  });

  viewports.forEach(view => {
    it(`should check font styles on ${view.name}`, () => {
      cy.viewport(view.width, view.height);
      cy.visit(websiteUrl);

      resultsByViewport[view.name] = [];

      // Use default selectors if style guide is not available
      const selectorsToCheck = isStyleGuideAvailable ? Object.keys(fontRules) : defaultSelectors;

      cy.wrap(selectorsToCheck).each(selector => {
        cy.document().then(doc => {
          const elements = doc.querySelectorAll(selector);
          if (elements.length === 0) {
            resultsByViewport[view.name].push({
              Selector: selector,
              Text: '',
              Status: isStyleGuideAvailable ? 'Not Found' : 'No Style Guide',
              Expected_fontSize: '-',
              Actual_fontSize: '-',
              Expected_fontWeight: '-',
              Actual_fontWeight: '-',
              Expected_fontFamily: '-',
              Actual_fontFamily: '-',
              Expected_lineHeight: '-',
              Actual_lineHeight: '-',
              Variant: '-'
            });
            cy.log(`Selector "${selector}" not found on viewport ${view.name}.`);
            return;
          }

          cy.wrap(elements).each(($el, index) => {
            cy.wrap($el).then($element => {
              const computedStyle = window.getComputedStyle($element[0]);
              const actual = {
                fontSize: computedStyle.fontSize,
                fontWeight: computedStyle.fontWeight,
                fontFamily: computedStyle.fontFamily,
                lineHeight: computedStyle.lineHeight,
              };

              if (isStyleGuideAvailable) {
                // If style guide exists, include comparison and status
                const rulesForSelector = fontRules[selector];
                rulesForSelector.forEach(expected => {
                  const columnName = viewportToColumnMap[view.name];
                  const expectedFontSize = expected[columnName] || '';
                  const expectedLineHeight = expected.lineHeight || '';

                  const isFontFamilyMatch = actual.fontFamily.includes(expected.fontFamily);
                  const isFontSizeMatch = normalizeFontSize(actual.fontSize) === normalizeFontSize(expectedFontSize);
                  const isFontWeightMatch = actual.fontWeight === expected.fontWeight;
                  const isLineHeightMatch = actual.lineHeight === expectedLineHeight;

                  let mismatchDetails = [];
                  if (!isFontFamilyMatch) mismatchDetails.push('Font Family');
                  if (!isFontSizeMatch) mismatchDetails.push('Font Size');
                  if (!isFontWeightMatch) mismatchDetails.push('Font Weight');
                  if (!isLineHeightMatch) mismatchDetails.push('Line Height');

                  const status = mismatchDetails.length === 0 ? 'Match' : `Mismatch: ${mismatchDetails.join(', ')}`;

                  resultsByViewport[view.name].push({
                    Selector: selector,
                    Variant: expected.variant || '',
                    Text: $element.text().trim().slice(0, 50),
                    Status: status,
                    Expected_fontSize: expectedFontSize,
                    Actual_fontSize: actual.fontSize,
                    Expected_fontWeight: expected.fontWeight || '',
                    Actual_fontWeight: actual.fontWeight,
                    Expected_fontFamily: expected.fontFamily || '',
                    Actual_fontFamily: actual.fontFamily,
                    Expected_lineHeight: expectedLineHeight || '',
                    Actual_lineHeight: actual.lineHeight
                  });
                });
              } else {
                // If style guide is missing, add "-" for expected values
                resultsByViewport[view.name].push({
                  Selector: selector,
                  Text: $element.text().trim().slice(0, 50),
                  Status: 'No Style Guide',
                  Expected_fontSize: '-',
                  Actual_fontSize: actual.fontSize,
                  Expected_fontWeight: '-',
                  Actual_fontWeight: actual.fontWeight,
                  Expected_fontFamily: '-',
                  Actual_fontFamily: actual.fontFamily,
                  Expected_lineHeight: '-',
                  Actual_lineHeight: actual.lineHeight,
                  Variant: '-'
                });
              }
            });
          });
        });
      });
    });
  });

  after(() => {
    const hasData = Object.values(resultsByViewport).some(data => 
      Array.isArray(data) && data.length > 0
    );

    if (!hasData) {
      console.log('No data to write to Excel file');
      return;
    }

    cy.task('writeExcelSheets', {
      data: resultsByViewport,
      filename: 'cypress/results/responsive-font-check-sheets.xlsx',
    });
  });
});
