//put custom drop down menu
function onOpen(e) {
  SpreadsheetApp.getUi() 
    .createMenu('Custom Utilities')
    .addItem('Trimmer', 'removeEmptyRowsActive')
    .addItem('Import Consultant Info', 'importTAInfo')
    .addItem('Import Client Info', 'importClientInfo')
    .addToUi();
}

//changes column index number to a letter 
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
    
  var keys =[],
      directory = []
  
  //hr Setup
  var hrWorkbookId = "1TKCxe3hugvi1d3UoWfbBXNW46S7puqafryx1bv230zI",
      hrSheetName = "Personnel",
      hrGrabbedInfo = [
        'Consultant',
        'Consultant Phone'
      ],
      hrPrimaryKey = 'Consultant',
      hrHeaderRowNumber = 1
    
  //client Setup
  var clientWorkbookId = "1TKCxe3hugvi1d3UoWfbBXNW46S7puqafryx1bv230zI",
      clientSheetName = "Locations",
      clientGrabbedInfo = [
        'Location',
        'Location Address',
        'On Site Contact',
        'On Site Contact Phone'
      ],
      clientPrimaryKey = 'Location',
      clientHeaderRowNumber = 1

//gets run first to config each type of import
//if more types are added insert another else if clause for that type
function setupImport(type) {
  var workbookId, sheetName, grabbedInfo, headerRowNumber
  if (type === 'hr') {
    workbookId = hrWorkbookId
    sheetName = hrSheetName
    grabbedInfo = hrGrabbedInfo
    headerRowNumber = hrHeaderRowNumber
  } else if (type === 'client') {
    workbookId = clientWorkbookId
    sheetName = clientSheetName
    grabbedInfo = clientGrabbedInfo
    headerRowNumber = clientHeaderRowNumber
  }
  var importedData = SpreadsheetApp.openById(workbookId).getSheetByName(sheetName),
      rangeData = importedData.getDataRange(),
      importedValues = rangeData.getValues()
  
  //this makes the directory from the source workbook to search the destination sheet for
  function makeObject() {
    importedValues.forEach( function (importedLine, ilIndex) {
      if (ilIndex === 0) {
        keys = importedLine
      } else {
        var tempObject = {}
        importedLine.forEach( function (dataPiece, dpIndex) {
          if (grabbedInfo.indexOf(keys[dpIndex]) > -1 && keys[dpIndex] === grabbedInfo[0]) {
            tempObject[keys[dpIndex]] = dataPiece
          } else if (grabbedInfo.indexOf(keys[dpIndex]) > -1) {
            tempObject[keys[dpIndex]] = {row: ilIndex, col: dpIndex}
          }
        })
        directory.push(tempObject)
      }
    })
    return directory
  }
directory = makeObject()
return {workbookId: workbookId, sheetName: sheetName, grabbedInfo: grabbedInfo, headerRowNumber: headerRowNumber}
}


//main top level function that inserts an =IMPORTRANGE to qualifying cells.
//if you add more types include another else if clause with it's primaryKey
function importInfo (type) {
  var config, primaryKey
  if (type === 'hr'){
    primaryKey = hrPrimaryKey
    config = setupImport(type)
  } else if (type === 'client') {
    primaryKey = clientPrimaryKey
    config = setupImport(type)
  } 
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      searchedValues = ss.getDataRange().getValues(),
      pkNamePosition
  searchedValues.forEach( function (row, rIndex) {
    row.forEach( function (col, cIndex) {
      if (col !== '') {
        directory.forEach( function ( dir, dIndex) {
          if (dir[primaryKey] === col ) {
            pkNamePosition = {row:rIndex, col:cIndex}
            var dirKeys = Object.keys(dir)
            dirKeys.forEach( function (key, keyIndex) {
            var dataIndexPosition = searchedValues[2].indexOf(key)
              if (dataIndexPosition > -1 && key !== primaryKey) {
                var cell = ss.getRange( pkNamePosition.row+1, dataIndexPosition + 1),
                    narrowRow = dir[key].row +1,
                    narrowCol =  columnToLetter(dir[key].col+1),
                    sourceCell = narrowCol + narrowRow,
                    narrowRange = SpreadsheetApp.openById(config.workbookId).getSheetByName(config.sheetName),
                    import = '=IMPORTRANGE("' + config.workbookId +  '", "' + config.sheetName  + '!' + sourceCell + ":" + sourceCell  + '")'
                cell.setFormula(import)
              }
            })
          }
        })
      }
    })
  })
}

function importTAInfo () {
  importInfo('hr')
}

function importClientInfo () {
  importInfo('client')
}

function removeEmptyRowsActive(){
  var sh = SpreadsheetApp.getActiveSheet()
  var maxRows = sh.getMaxRows()
  var lastRow = sh.getLastRow()
  if (maxRows !== lastRow) {
    sh.deleteRows(lastRow+1, maxRows-lastRow)
  }
}