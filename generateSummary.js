const Excel = require('exceljs/modern.nodejs');
const _ = require('lodash');

main();

function main() {
  const clArgs = getArgs();
  let filename = clArgs[0];
  let allDataArr = [];
  let inputWorkbook = new Excel.Workbook();
  let outputWorkbook = new Excel.Workbook();

  let promises = [];
  console.log("Reading file:", filename, '\n');
  readExcelFile(inputWorkbook, filename)
    .then(function() {
      let sheetNames = getWorksheetNames(inputWorkbook);
      _.each(sheetNames, function(inputSheet) {
        let worksheet = getExcelSheet(inputWorkbook, inputSheet);
        let sheetObj = getWorksheetObj(worksheet, 1, 4);
        allDataArr.push(sheetObj);
        let sheetName = inputSheet + ' SUMMARY';
        addExcelSheetToWorkbook(outputWorkbook, sheetObj, sheetName);
        console.log(sheetName + 'added to workbook.')
      });
      let mergedTotalObj = createTotalDataObj(allDataArr);
      addExcelSheetToWorkbook(outputWorkbook, mergedTotalObj, "SUMMARY TOTAL");
      console.log('SUMMARY TOTAL added to workbook.\n')
      let outputFileName = 'Retail_Summary.xlsx'
      outputWorkbook.xlsx.writeFile(outputFileName)
        .then(function() {
          console.log('Output written to', outputFileName);

          console.log('\nDONE');
        });
    });
}

function getArgs() {
  return process.argv.slice(2);
}

function readExcelFile(workbook, filename) {
  return workbook.xlsx.readFile(filename);
}

function getWorksheetNames(workbook) {
  return _.map(workbook.worksheets, function(sheetObj) {
    return sheetObj.name;
  });
}

function getExcelSheet(workbook, worksheetName) {
  let sheet = _.find(workbook.worksheets, function(sheetObj) {
        return sheetObj.name === worksheetName;
      });
  if(sheet) {
    return sheet;
  }
  console.log('Sheet not found');
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
  let priceCol = "H";
  let price = row.getCell(priceCol);
  let startColIdx = 13; // TODO: dont hardcode this
  for(var idx = startColIdx; idx < startColIdx+12; idx++) {
    let val = row.getCell(idx).value;
    if(val !== null) {
      val = val.result !== undefined ? val.result : val;
    } else {
      val = 0;
    }
    let revIdx = idx-startColIdx;
    if(franchiseObj.revenues[revIdx] == undefined) {
      franchiseObj.revenues[revIdx] = 0;
    }
    franchiseObj.revenues[revIdx] += (price*val);
  }
  // TODO: Total up revenues in each category in this function
  // console.log(franchiseObj);
}

function addExcelSheetToWorkbook(workbook, sheetObj, sheetName) {
  // console.log(sheetObj);
  let sheet = workbook.addWorksheet(sheetName);
    // TODO: get these from input file and make dynamic
  let colHeaders = ["products", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
  sheet.columns = getSheetColumns(colHeaders);
  _sumCategories(sheetObj);
  _addCategories(sheet, sheetObj, colHeaders);

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
    } else {
      let revId = idx-1; // TODO: make the rev header keys months, not idx
      rowObj[header] = dataObj.revenues[revId];
    }
  });
  return rowObj;
}
function createTotalDataObj(allDataArr) {
  // mergeWith mutates object
  return _.reduce(allDataArr, function(acc, curr) {
    return _.mergeWith(acc, curr, mergeWithHelper);
  });
}

function mergeWithHelper(objVal, srcVal, key) {
  if(key === "revenues") {
    for(let month in objVal) {
      if(srcVal[month] === undefined) {
        srcVal[month] = 0;
      }
      objVal[month] += srcVal[month];
    }
    return objVal;
  }
}
