const Excel = require('exceljs/modern.nodejs');
const _ = require('lodash');

main();

function main() {
  const clArgs = getArgs();
  let filename = clArgs[0];
  let inputSheets = [".COM US", "b&m us"];
  let allDataArr = [];
  // let inputSheets = [".COM US"];
  let outputWorkbook = new Excel.Workbook();
  let promises = [];
  _.each(inputSheets, function(inputSheet) {
    let promise = getExcel(filename, inputSheet)
      .then(function(worksheet) {
        let sheetObj = getWorksheetObj(worksheet, 1, 4);
        allDataArr.push(sheetObj);
        let sheetName = inputSheet + ' SUMMARY';
        addExcelSheetToWorkbook(outputWorkbook, sheetObj, sheetName);
      });
      promises.push(promise);
  });
  Promise.all(promises)
    .then(function(values) {
      outputWorkbook.xlsx.writeFile('Retail_Summary.xlsx')
        .then(function() {
          console.log('DONE');
        });
    });
}

function getArgs() {
  return process.argv.slice(2);
}

function getExcel(filename, worksheetName) {
  var workbook = new Excel.Workbook();
  return workbook.xlsx.readFile(filename)
  .then(function() {
    let worksheet;
    if(worksheetName) {
      worksheet = _.find(workbook.worksheets, function(sheetObj) {
        return sheetObj.name === worksheetName;
      });
    }
    return worksheet ? worksheet : workbook;
  });
}

function getWorksheetObj(worksheet, colStartIdx, rowStartIdx) {
  let rows = worksheet._rows;
  let dataObj = {};

  for(var idx = rowStartIdx; idx < rows.length-1; idx++) {
    let row = rows[idx];
    let category = row.getCell(colStartIdx+1).value; // cols start at 1, not 0
    if(!category) {
      continue;
    }
    if(dataObj[category] === undefined) {
      dataObj[category] = {
        name: category,
        franchises: {}
      };
    }
    let franchise = row.getCell(colStartIdx+2).value;
    // TODO: createFranchiseObj function
    if(dataObj[category].franchises[franchise] === undefined) {
      dataObj[category].franchises[franchise] = {
        name: franchise,
        revenues: {}
      };
    }
    addRevenuesToFranchise(dataObj[category].franchises[franchise], row);
  }
  return dataObj;
}

function addRevenuesToFranchise(franchiseObj, row) {
  let startColIdx = 37+1;
  for(var idx = startColIdx; idx < startColIdx+12; idx++) {
    let val = row.getCell(idx).value;
    if(val !== null) {
      val = val.result !== undefined ? val.result : val;
    } else {
      val = 0;
    }
    let revIdx = idx-38;
    if(franchiseObj.revenues[revIdx] == undefined) {
      franchiseObj.revenues[revIdx] = 0;
    }
    franchiseObj.revenues[revIdx] += val;
  }
  // TODO: Total up revenues in each category in this function
  // console.log(franchiseObj);
}

function addExcelSheetToWorkbook(workbook, sheetObj, sheetName) {
  // console.log(sheetObj);
  let sheet = workbook.addWorksheet(sheetObj.inputSheetName);
  let colHeaders = ["products", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
  sheet.columns = getSheetColumns(colHeaders);
  _sumCategories(sheetObj);
  _addCategories(sheet, sheetObj, colHeaders);
  // TODO: get these from input file
  // console.log(sheet.columns);
}

function getSheetColumns(headers) {
  let cols = _.map(headers, function(header) {
    return { header: header, key: header, width: 15};
  });
  cols[0].header = '';
  cols[0].width = 40;
  return cols;
}

function _sumCategories(sheetObj) {
  for(let category in sheetObj) {
    sheetObj[category].revenues = {};
    let franchises = sheetObj[category].franchises;
    for(let franchise in franchises) {
      for(let month in franchises[franchise].revenues) {
        if(sheetObj[category].revenues[month] == undefined) {
          sheetObj[category].revenues[month] = 0;
        }
        sheetObj[category].revenues[month] += franchises[franchise].revenues[month];
      }
    }
  }
}

function _addCategories(sheet, sheetObj, colHeaders) {
  for(let category in sheetObj) {
    let categoryObj = sheetObj[category];
    let categoryRow = _createRowObj(categoryObj, colHeaders);
      sheet.addRow(categoryRow);
      sheet.getRow(sheet.rowCount).fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{argb:'D9E1F2'},
      };
      // console.log(sheet.rowCount);
      for(let franchise in categoryObj.franchises) {
        let franchiseObj = categoryObj.franchises[franchise];
        let franchiseRow = _createRowObj(franchiseObj, colHeaders, true);
        sheet.addRow(franchiseRow);
        let currRow = sheet.getRow(sheet.rowCount);
        currRow.getCell(1).alignment = { indent: 1 };
      }
  }
}

function _createRowObj(dataObj, colHeaders, isTotalRow) {
  let rowObj = {};
  _.each(colHeaders, function(header, idx) {
    if(idx === 0) {
      rowObj[header] = dataObj.name;
      // if(isTotalRow) {
      //   console.log('alignment');
      //   rowObj.style = {
      //     fill: {
      //       bgColor:{argb:'FF0000FF'}
      //     }
      //   }
      // }
    } else {
      let revId = idx-1; // TODO: make the rev header keys months, not idx
      rowObj[header] = dataObj.revenues[revId];
    }
  });
  return rowObj;
}
