const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();

const Excel = require('exceljs');
const workbook = new Excel.Workbook();

function getCellResult(worksheet, cellLabel) {
  return new Promise((resolve, reject) => {
    try {
      resolve(parser.parse(cellLabel));
      if (worksheet.getCell(cellLabel).formula) {
        resolve(parser.parse(worksheet.getCell(cellLabel).formula));
      } else {
        resolve(worksheet.getCell(cellCoord.label).value);
      }
    } catch {
      reject('error');
    }
  });
}

workbook.xlsx.readFile('./easy.xlsx').then(async () => {
  var worksheet = workbook.getWorksheet(1);

  parser.on('callCellValue', function(cellCoord, done) {
    if (worksheet.getCell(cellCoord.label).formula) {
      done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
    } else {
      done(worksheet.getCell(cellCoord.label).value);
    }
  });

  parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
    var fragment = [];

    for (var row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
      var colFragment = [];

      for (var col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
        colFragment.push(worksheet.getRow(row + 1).getCell(col + 1).value);
      }

      fragment.push(colFragment);
    }

    if (fragment) {
      done(fragment);
    }
  });

  worksheet.getCell('A2').value = 100;
  var result = await getCellResult(worksheet, 'A$1&" world"');
  console.log(result);
});
