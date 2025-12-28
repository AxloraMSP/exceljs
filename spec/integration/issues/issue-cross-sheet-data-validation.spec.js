const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/cross-sheet-data-validation.test.xlsx';

describe('github issues', () => {
  describe('cross-sheet data validation', () => {
    it('should preserve cross-sheet reference when reading (simple sheet name)', () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Sheet1');
      const ws2 = wb.addWorksheet('Sheet2');

      ws2.getCell('A1').value = 'Option1';
      ws2.getCell('A2').value = 'Option2';
      ws2.getCell('A3').value = 'Option3';

      ws1.getCell('A1').dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['Sheet2!$A$1:$A$3'],
        showErrorMessage: true,
        error: 'Please select a valid option',
        errorTitle: 'Invalid Selection',
      };

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(wb2 => {
          const ws1Read = wb2.getWorksheet('Sheet1');
          const dataValidation = ws1Read.getCell('A1').dataValidation;

          expect(dataValidation).to.not.be.undefined();
          expect(dataValidation.type).to.equal('list');
          expect(dataValidation.formulae).to.be.an('array');
          expect(dataValidation.formulae.length).to.equal(1);
          expect(dataValidation.formulae[0]).to.equal('Sheet2!$A$1:$A$3');
        });
    });

    it('should preserve cross-sheet reference with quoted sheet name', () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Sheet1');
      const ws2 = wb.addWorksheet('Sheet 2');

      ws2.getCell('A1').value = 'Value1';
      ws2.getCell('A2').value = 'Value2';

      ws1.getCell('B1').dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ["'Sheet 2'!$A$1:$A$2"],
        showErrorMessage: true,
        error: 'Please select a valid value',
      };

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(wb2 => {
          const ws1Read = wb2.getWorksheet('Sheet1');
          const dataValidation = ws1Read.getCell('B1').dataValidation;

          expect(dataValidation).to.not.be.undefined();
          expect(dataValidation.formulae).to.be.an('array');
          expect(dataValidation.formulae.length).to.equal(1);
          expect(dataValidation.formulae[0]).to.include('Sheet 2');
          expect(dataValidation.formulae[0]).to.include('!');
        });
    });

    it('should preserve cross-sheet reference for custom validation type', () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Sheet1');
      const ws2 = wb.addWorksheet('DataSheet');

      ws2.getCell('B1').value = 10;
      ws2.getCell('B2').value = 20;

      ws1.getCell('C1').dataValidation = {
        type: 'custom',
        allowBlank: true,
        formulae: ['DataSheet!$B$1+DataSheet!$B$2'],
        showErrorMessage: true,
        error: 'Invalid value',
      };

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(wb2 => {
          const ws1Read = wb2.getWorksheet('Sheet1');
          const dataValidation = ws1Read.getCell('C1').dataValidation;

          expect(dataValidation).to.not.be.undefined();
          expect(dataValidation.type).to.equal('custom');
          expect(dataValidation.formulae).to.be.an('array');
          expect(dataValidation.formulae.length).to.equal(1);
          expect(dataValidation.formulae[0]).to.include('DataSheet');
          expect(dataValidation.formulae[0]).to.include('!');
        });
    });

    it('should not convert cross-sheet reference to number for whole/textLength/decimal types', () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Sheet1');
      const ws2 = wb.addWorksheet('Sheet2');

      ws2.getCell('A1').value = 5;

      ws1.getCell('D1').dataValidation = {
        type: 'whole',
        operator: 'equal',
        formulae: ['Sheet2!$A$1'],
        showErrorMessage: true,
        error: 'Value must match',
      };

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(wb2 => {
          const ws1Read = wb2.getWorksheet('Sheet1');
          const dataValidation = ws1Read.getCell('D1').dataValidation;

          expect(dataValidation).to.not.be.undefined();
          expect(dataValidation.type).to.equal('whole');
          expect(dataValidation.formulae).to.be.an('array');
          expect(dataValidation.formulae.length).to.equal(1);
          expect(dataValidation.formulae[0]).to.be.a('string');
          expect(dataValidation.formulae[0]).to.equal('Sheet2!$A$1');
        });
    });
  });
});

