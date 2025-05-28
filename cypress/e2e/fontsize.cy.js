/// <reference types="cypress" />
const XLSX = require('xlsx');
const fs = require('fs');

const fontRules = {
  h1: [
    //Large H1
    { fontSize: '48px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
    // Normal H1
    { fontSize: '44px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
    // Small H1
    { fontSize: '42px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

  h2: [
     //Large H2
    { fontSize: '34px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
    // Normal H2
    { fontSize: '33px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

  h3:[
    { fontSize: '24px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

h4:[
    { fontSize: '21px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

h5:[
    { fontSize: '19px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

h6:[
    { fontSize: '17px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

p:[
    //Large P
    { fontSize: '21px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
    //Normal P
    { fontSize: '17px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

Subtitle:[
     { fontSize: '14px', fontWeight: '800', fontFamily: 'Inter' }, //Add lineHeight: '' , if provied in the style guide
  ],
Button:[
     { fontSize: '16px', fontWeight: '800', fontFamily: 'Inter, sans-serif' }, //Add lineHeight: '' , if provied in the style guide
  ],

};


describe('Font Style Checker', () => {
  before(() => {
    cy.wrap([]).as('results');
  });

  it('should check font styles as per design system', function () {
    cy.viewport(1920, 1080);
    cy.visit('https://49126198.hs-sites-na2.com/');

    cy.document().then((doc) => {
      Object.keys(fontRules).forEach((selector) => {
        const elements = doc.querySelectorAll(selector);

        if (elements.length === 0) {
          // If no elements, add "Ignored" row
          const result = {
            Selector: selector,
            Index: 'N/A',
            Text: 'N/A',
            Status: 'Ignored',
            Expected_fontSize: fontRules[selector][0].fontSize,
            Expected_fontWeight: fontRules[selector][0].fontWeight,
            Expected_fontFamily: fontRules[selector][0].fontFamily,
            Actual_fontSize: 'N/A',
            Actual_fontWeight: 'N/A',
            Actual_fontFamily: 'N/A',

          };

          cy.get('@results').then((results) => {
            results.push(result);
            cy.wrap(results).as('results');
          });
          return;
        }

        // Only check the first element for each selector
        const el = elements[0];
        const computedStyle = window.getComputedStyle(el);

        const actual = {
          fontSize: computedStyle.fontSize,
          fontWeight: computedStyle.fontWeight,
          fontFamily: computedStyle.fontFamily,
        };

        let matchFound = false;
        for (let rule of fontRules[selector]) {
          const isFontSizeMatch = actual.fontSize === rule.fontSize;
          const isFontWeightMatch = actual.fontWeight === rule.fontWeight;
          const isFontFamilyMatch = actual.fontFamily === rule.fontFamily;

          if (isFontSizeMatch && isFontWeightMatch && isFontFamilyMatch) {
            matchFound = true;
            break;
          }
        }

        const result = {
          Selector: selector,
          Index: 1,
          Text: el.textContent.trim().slice(0, 50),
          Status: matchFound ? 'Match' : 'Mismatch',
          Actual_fontSize: actual.fontSize,
          Actual_fontWeight: actual.fontWeight,
          Actual_fontFamily: actual.fontFamily,
          Expected_fontSize: fontRules[selector][0].fontSize,
          Expected_fontWeight: fontRules[selector][0].fontWeight,
          Expected_fontFamily: fontRules[selector][0].fontFamily,
        };

        cy.get('@results').then((results) => {
          results.push(result);
          cy.wrap(results).as('results');
        });
      });
    });
  });

  after(function () {
    cy.get('@results').then((results) => {
      const filename = 'font-results.xlsx';
      cy.task('writeExcel', { data: results, filename });
      cy.log(`Results saved to fixtures/${filename}`);
    });
  });
});
