describe('Responsive Font Style Checker with Variants and Sheets', () => {
  let fontRules = {};
  let resultsByViewport = {};
  let websiteUrl = '';

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

  before(() => {
    cy.task('readExcel', {
      filePath: './cypress/fixtures/Style Guide.xlsx',
    }).then(data => {
      fontRules = {};
      
      // Read the homepage URL from urls.json
      cy.fixture('urls.json').then(urls => {
        websiteUrl = urls.homepage;
        cy.log('Website URL:', websiteUrl);
      });
      
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
          // SmallMobile: row[-]?.trim(),
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
      console.log('Website URL:', websiteUrl);
    });
  });

  viewports.forEach(view => {
    it(`should check font styles on ${view.name}`, () => {
      cy.viewport(view.width, view.height);
      cy.visit(websiteUrl);

      resultsByViewport[view.name] = [];

      const selectors = Object.keys(fontRules);

      cy.wrap(selectors).each(selector => {
        cy.document().then(doc => {
          const elements = doc.querySelectorAll(selector);
          if (elements.length === 0) {
            // Add a record for not found selectors
            resultsByViewport[view.name].push({
              Selector: selector,
              Status: 'Not Found',
              Text: '',
              Expected_fontSize: '',
              Actual_fontSize: '',
              Actual_lineHeight: '',
              Expected_lineHeight: '',
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
              cy.wrap($el).then($element => {
                const computedStyle = window.getComputedStyle($element[0]);

                const actual = {
                  fontSize: computedStyle.fontSize,
                  lineHeight: computedStyle.lineHeight,
                  fontWeight: computedStyle.fontWeight,
                  fontFamily: computedStyle.fontFamily,
                };

                const columnName = viewportToColumnMap[view.name];
                const expectedFontSize = expected[columnName] || '';

                const isFontFamilyMatch = actual.fontFamily.includes(expected.fontFamily);
                const isFontSizeMatch = normalizeFontSize(actual.fontSize) === normalizeFontSize(expectedFontSize);
                const isFontWeightMatch = actual.fontWeight.toString() === expected.fontWeight?.toString();
                const isLineHeightMatch = normalizeFontSize(actual.lineHeight) === normalizeFontSize(expected.lineHeight);

                let mismatchDetails = [];
                if (!isFontFamilyMatch) mismatchDetails.push('Font Family');
                if (!isFontSizeMatch) mismatchDetails.push('Font Size');
                if (!isLineHeightMatch) mismatchDetails.push('Line Height');
                if (!isFontWeightMatch && expected.fontWeight) mismatchDetails.push('Font Weight');

                const status = mismatchDetails.length === 0 ? 'Match' : `Mismatch: ${mismatchDetails.join(', ')}`;
                
                resultsByViewport[view.name].push({
                  Selector: selector,
                  Variant: expected.variant || '',
                  Text: $element.text().trim().slice(0, 50),
                  Status: status,
                  Expected_fontSize: expectedFontSize,
                  Actual_fontSize: actual.fontSize,
                  Expected_lineHeight: expected.lineHeight || '',
                  Actual_lineHeight: actual.lineHeight,
                  Expected_fontWeight: expected.fontWeight || '',
                  Actual_fontWeight: actual.fontWeight,
                  Expected_fontFamily: expected.fontFamily || '',
                  Actual_fontFamily: actual.fontFamily
                });
              });
            });
          });
        });
      });
    });
  });

  after(() => {
    // Ensure we have valid data to write
    const hasData = Object.values(resultsByViewport).some(data => 
      Array.isArray(data) && data.length > 0
    );

    if (!hasData) {
      console.log('No data to write to Excel file');
      return;
    }

    cy.task('writeExcelSheets', {
      data: resultsByViewport,
      filename: './cypress/results/responsive-font-check-sheets.xlsx',
    });
  });
});
