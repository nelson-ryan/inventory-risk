//AlaSQLGS: SQL capabilities within Apps Script
//const alasql = AlaSQLGS.load();


function mergePastDue() {
  mergeItUp("Past Due")
}
function mergeRiskRevenue() {
  mergeItUp("Risk")
}

function mergeItUp(category) {

  const datafile = SpreadsheetApp.openById("datafiledatafiledatafiledatafiledatafile")
  const sheetParts = getSheetById(datafile, 824824824)
  const sheetOutlines = getSheetById(datafile, 321321321)
  const sheetBOM = getSheetById(datafile, 075075075)

  const outputfile = SpreadsheetApp.openById("outputfileoutputfileoutputfileoutputfile")
  
  //sheetRisk = getSheetById(outputfile, 1234787403)
  const sheetRisk = ((category == 'Past Due' ?
                      getSheetById(outputfile, 981981981) :
                      getSheetById(outputfile, 123123123)))

  let parts = createDataObject(sheetParts)
  let outlines = createDataObject(sheetOutlines)
  let report = createDataObject(sheetRisk)

  // report has the right object structure, but clear arrays by replacing with empty
  for(i in report) {
    if (i != 'headers') {report[i] = []}} // except headers; keep those!

  // if output report has no header, use combined (unique) headers from each source
  if (report.headers == []) {
    report.headers = report.headers.concat(outlines.headers.filter(header => !report.headers.includes(header)))
    report.headers = report.headers.concat(parts.headers.filter(header => !report.headers.includes(header)))
  }


  // filter out "green" parts
  let partsStatusColumn = parts.headers.indexOf("Status")
  let statusIndices = parts.data.map(row => row[partsStatusColumn] != 'Green')
  for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds']) {
    parts[key] = parts[key].filter((row, i) => (statusIndices[i]))
  }
  // filter out Risk vs Past Due outlines
  let outLinecategoryColumn = outlines.headers.indexOf("Risk Level")
  if (category == "Past Due") {
    statusIndices = outlines.data.map(row => row[outLinecategoryColumn] == "Past Due")
  } else {
    statusIndices = outlines.data.map(row => row[outLinecategoryColumn] != "Past Due")
  }
  for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds']) {
    outlines[key] = outlines[key].filter((row, i) => (statusIndices[i]))
  }

  // SORTING OF OUTLINES - must be done in Outlines data sheet (may be for the better, as it will allow for any kind of sorting)
  // // sort outlines descending by Revenue total
  // let sortColumn = outlines.headers.indexOf("Revenue")
  // outlines.data = outlines.data.sort(function (a, b) {
  //                                       if (a[sortColumn] === b[sortColumn]) {return 0}
  //                                       else {return ((a[sortColumn] > b[sortColumn]) ? -1 : 1)}
  //                                     })
  
  // get indices for column sorting
  for (i of [parts, outlines]) {
    i.indices = i.headers.map(header => report.headers.indexOf(header))
    }
  
  //// BOM mapping
  // index of part number in each source
  var parentOL = outlines.headers.indexOf(datafile.getRangeByName("OutlineHeader_MaterialNo").getValue())
  var gatingpartno = parts.headers.indexOf(datafile.getRangeByName("OutlineHeader_MaterialNo").getValue())

  var bomList = sheetBOM.getDataRange().getValues()
  
  bomList = getBomMapping(bomList,
                          outlineColName = datafile.getRangeByName("BOMHeader_StartAssembly").getValue(),
                          partColName = datafile.getRangeByName("BOMHeader_Part").getValue(),
                          qtyColName = datafile.getRangeByName("BOMHeader_QtyPerNHA").getValue())

  
  var usagePerIndex = report.headers.indexOf(outputfile.getRangeByName("RiskRevHeader_UsagePer").getValue())
  //add each OL and part in turn
  for (o in outlines.data) {
      var keep = false
      // add outline line
      for (key of ['data', 'fontweights', 'fontcolors']) {  
        var temprow = new Array(report.headers.length)
        for (i in outlines.indices) {temprow[outlines.indices[i]] = outlines[key][o][i]}
        report[key].push(temprow)
      }

      //hacky step for background colors, fill with grey for outline rows
      var temprow = new Array(report.headers.length)
      for (i in outlines.indices) {temprow[outlines.indices[i]] = outlines.backgrounds[o][i]}
      temprow = Array.from(temprow, b => ((b == undefined || b == "#ffffff") ? "#eeeeee" : b))
      report.backgrounds.push(temprow)

      // get outline number, part list from BOM
      var outlineNo = outlines.data[o][parentOL].trim()
      //check it's even in the BOM list; if not, probably need to update/paste in from SAP
      if(outlineNo in bomList) {
        /// creating a new 2D array of parts to concatenate under outline will allow for sorting of parts within outline
        var bomOutline = bomList[outlineNo]
        //for (component of componentList) {
          for (var p in parts.data) {
            if (parts.data[p][gatingpartno] in bomOutline) {
              for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds'])  {
                temprow = new Array(report.headers.length)
                for (i in parts.indices) {temprow[parts.indices[i]] = parts[key][p][i]}
                // assign Qty from BOM list to 'Usage Per' field
                if (key == 'data') {
                  var datatemprow = Array.from(temprow, (c,ind) => (ind == usagePerIndex ? bomOutline[parts.data[p][gatingpartno]] : c))
                  report['data'].push(datatemprow)
                  keep = true
                } else {report[key].push(temprow)}
              }
            }
          }
      }
      if (!keep) {
        for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds']) {  
        report[key].pop()
        }
      }
    }
  if (report.data.length <= 0) {return}
  // Pasting in sheet and formatting
  var firstrow = 1
  var firstcol = 1
  sheetRisk.getRange(row = firstrow, column = firstcol, numRows = sheetRisk.getMaxRows(), numColumns = sheetRisk.getMaxColumns())
            .clearContent()
  sheetRisk.getRange(row = firstrow, column = firstcol, numRows = sheetRisk.getMaxRows(), numColumns = sheetRisk.getMaxColumns())
            .clearFormat()
  var headerrow = sheetRisk.getRange(row = firstrow, column = firstcol, numRows = 1, numColumns = report.headers.length)
  headerrow.setValues([report.headers])
  
  headerrow.setBackground("#c7c7c7")
  headerrow.setFontWeight("bold")
  headerrow.setBorder(top = true, left = true, bottom = true, right = true, vertical = true,
                      horizontal = true, color = '#000000', style = SpreadsheetApp.BorderStyle.SOLID)
  headerrow.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)

  
  var outputrange = sheetRisk.getRange(row = firstrow + 1,
                       column = firstcol,
                       numRows = report.data.length,
                       numColumns = report.data[0].length)
  outputrange.setValues(report.data)

  outputrange.setBackgrounds(report.backgrounds)
  outputrange.setFontWeights(report.fontweights)
  outputrange.setFontColors(report.fontcolors)
  outputrange.setBorder(top = true, left = true, bottom = true, right = true, vertical = false,
                        horizontal = false, color = '#000000', style = SpreadsheetApp.BorderStyle.SOLID)
  outputrange.setBorder(top = null, left = null, bottom = null, right = null, vertical = null,
                        horizontal = true, color = '#c3c3c3', style = SpreadsheetApp.BorderStyle.SOLID)
  
  //// This sets the first column with vertical borders  
  // var partsstart = Math.max.apply(Math, outlines.indices) - 1 
  var partsstart = report.headers.indexOf("Usage / per") - 1
  
  var partsrange = sheetRisk.getRange(outputrange.getRow(), outputrange.getColumn() + partsstart,
                                      outputrange.getNumRows(), outputrange.getNumColumns() - partsstart)
  partsrange.setBorder(top = null, left = null, bottom = null, right = null, vertical = true,
                        horizontal = null, color = '#000000', style = SpreadsheetApp.BorderStyle.SOLID)
  
  // Outline rows
  let outlinerows = []
  let idColumnIndex = report.headers.indexOf("Revenue")
  for (var i in report.data) {
    if (report.data[i][idColumnIndex] != null && report.data[i][idColumnIndex] != '') {
      var rg = sheetRisk.getRange(outputrange.getRow() + parseInt(i), outputrange.getColumn(), numRows = 1, numColumns = outputrange.getNumColumns())
      outlinerows.push(rg)
    }
  }
  
  for (i of outlinerows) {
    i.setBorder(top = true, left = true, bottom = null, right = true, vertical = null,
                      horizontal = false, color = '#000000', style = SpreadsheetApp.BorderStyle.SOLID)
  }

  let formatColumns = sheetRisk.getRange(row = headerrow.getRow(), column = headerrow.getColumn(),
                                           numRows = sheetRisk.getMaxRows(),
                                           numColumns = report.headers.indexOf(outputfile.getRangeByName("RiskRevHeader_Action").getValue()))
  formatColumns.setHorizontalAlignment('center')
  
  formatColumns = sheetRisk.getRange(row = headerrow.getRow(), column = report.headers.indexOf(outputfile.getRangeByName("RiskRevHeader_Action").getValue()) + 1,
                                           numRows = sheetRisk.getMaxRows(),
                                           numColumns = 1)
  formatColumns.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  formatColumns.setHorizontalAlignment('left')
}


//wb = SpreadsheetApp.getActiveSpreadsheet()
function getSheetById(spreadsheet, sheetgid) {
  const sheets = spreadsheet.getSheets();
  for (var i in sheets) {
    if(spreadsheet.getSheets()[i].getSheetId()==sheetgid){
      return spreadsheet.getSheets()[i]
    }
  }
}


function createDataObject(sourcesheet) {
  var range = sourcesheet.getDataRange()
  var data = sourcesheet.getDataRange().getValues()
  var empty = (true)
  for (var i of data[0]) {
    if (i != '') {empty = false;
                  break}
  }
  let obj = {'headers' : [],
             'data': data,
             'backgrounds' : range.getBackgrounds(),
             'fontcolors' : range.getFontColors(),
             'fontweights': range.getFontWeights(),
             'indices' : []
             }
  if (!empty) {
        obj.headers = obj.data.shift()
        obj.backgrounds.shift()
        obj.fontcolors.shift()
        obj.fontweights.shift()
    } else {
      for (i in obj) obj[i] = [];
    }
  return obj
}


function testbommapping() {
  const datafile = SpreadsheetApp.openById("datafiledatafiledatafiledatafiledatafile")
  const sheetBOM = getSheetById(datafile, 075075075)
  var bom = sheetBOM.getDataRange().getValues()
  console.log(getBomMapping(bom, 'Start Assembly', 'Part', datafile.getRangeByName("BOMHeader_QtyPerNHA").getValue()))

}

function getBomMapping(bomArray, outlineColName, partColName, qtyColName) {
  
  let outlinecol = bomArray[0].indexOf(outlineColName)
  let partcol = bomArray[0].indexOf(partColName)
  let qtycol = bomArray[0].indexOf(qtyColName)
  // bomArray = bomArray.map( function (row) {return [row[outlinecol], row[partcol], row[qtycol]]})
  bomArray.shift()
  // filter out top level
  bomArray = bomArray.filter(r => r[outlinecol] != r[partcol])
  // Create object with Outline indices to part numbers
  let obj = {}
  for (var row of bomArray) {
    if (obj[row[outlinecol]] == undefined) {
      // add if not there
      obj[row[outlinecol]] = {}
      // filter full bomArray for this outline's parts
      var b = bomArray.filter(r => r[outlinecol] == row[outlinecol])
      // add total qty of parts for outline
      for (var line of b) {
        if (obj[row[outlinecol]][line[partcol]] == undefined) { // for each, part, check if not already added
          // then filter for that part (within the already-filtered outline list) and sum qtys
          obj[row[outlinecol]][line[partcol]] = b.filter(part => part[partcol] == line[partcol])
                                                  .map(part => part[qtycol])
                                                  .reduce((acc, qty) => acc + qty)
        }
      }
    }
  }
  return obj
}


function bomMapTest() {
  bomList = sheetBOM.getDataRange().getValues()
  bomList = getBomMapping(bomList, "Start Assembly", "Part")
  console.log('9999999-100' in bomList)
  console.log((bomList['9999999-100']))
}



function identifyDuplicates(arr) {
  var newData = []
  var indices  = []
  for (var i in arr) {
    var row = arr[i]
    var duplicate = false
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true
        break
      }
    }
    if (!duplicate) {
      newData.push(row)
      indices.push(parseInt(i))
    }
  }
  return indices
}



function copyitout() {
  const tierboard = SpreadsheetApp.openById("outputfileoutputfileoutputfileoutputfile")
  const sheetRiskTB = getSheetById(tierboard, 123123123)
  dataCopy(sheetRiskTB, "Outline")
  dataCopy(sheetRiskTB, "Part")
  //dataCopy(sheetPastDue, "Outline") //;  dataCopy(sheetPastDue, "Part")
}


// function to copy data over from the Risk and Past Due tables
function dataCopy(sourcesheet, filtertype) {
  // just use sandbox tab for testing
  const datafile = SpreadsheetApp.openById("datafiledatafiledatafiledatafiledatafile")
  
  // use filter type to identify target spreadsheet tab
  var destsheet = (filtertype == "Outline" ? getSheetById(datafile, 321321321) :
                  (filtertype == "Part"    ? getSheetById(datafile, 824824824) : null))
  if (!destsheet) {return null}

  var neededColumns = destsheet.getDataRange().getValues()[0]

  // var data = sourcesheet.getDataRange().getValues()
  var data = createDataObject(sourcesheet)

  // Pare down data until identifying and removing top column, using "Parent Outline" column as the indicator,
  // Needed if there is anything (e.g. summary table) above; BUT the rest of the code does NOT allow for this
  var parentOL = data.headers.indexOf("Type")
  while (parentOL == -1) {
    data.headers = data.data.shift()
    data.backgrounds.shift()
    data.fontcolors.shift()
    data.fontweights.shift()
    parentOL = data.headers.indexOf("Type")
  } 
  
  // filter for rows with the given type value
  var filterindex = data.data.map(row => row[parentOL] == filtertype)
  for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds']) {
    data[key] = data[key].filter((_,i) => filterindex[i])
  }

  // reduce to just relevant columns
  // get indices first
  var indices = neededColumns.filter(col => data.headers.includes(col))
                              .map(col => data.headers.indexOf(col))
  
  // Function modified from stackoverflow.com/questions/62565632/
  // underscore is item itself, and i is the index (second parameter of filter's callback)
  data.headers = data.headers.filter((_,i) => indices.includes(i))
  // apply the same filter to all rows
  for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds']) {
    data[key] = data[key].map(row => row.filter((_,i) => indices.includes(i)))
  }
  
  var duplicateindices = identifyDuplicates(data.data)
  for (key of ['data', 'fontweights', 'fontcolors', 'backgrounds']) {
    data[key] = data[key].filter((_,i) => duplicateindices.includes(i))
  }


  let destsheetfullrange = destsheet.getRange(row = 1, column = 1, numRows = destsheet.getMaxRows(), numColumns = destsheet.getMaxColumns())
  destsheetfullrange.clearContent()
  destsheetfullrange.clearFormat()

  // // paste values into corresponding tab in the Data spreadsheet, //// after existing data
  destsheet.getRange(row = 1, column = 1, numRows = 1, numColumns = data.headers.length
                     ).setValues([data.headers])
  var pasterange = destsheet.getRange(row = 2,                                     //row = destsheet.getDataRange().getLastRow() + 1,
                    column = 1,//
                    numRows = data.data.length,
                    numColumns = data.data[0].length
                    )
  pasterange.setValues(data.data)
  pasterange.setBackgrounds(data.backgrounds)
  pasterange.setFontWeights(data.fontweights)
  pasterange.setFontColors(data.fontcolors)
}

// BELOW ARE NOT USED

// insert function is completely unnecessary as written, as the formula will just be in the sheet
// However, it may well still prove useful, adapting it for pulling part counts from BOM sheet
function insertLookupFormula() {
  var allRange = sheetRisk.getDataRange()
  
  var headerRange = sheetRisk.getRange(row = allRange.getRow(),
                                         column = allRange.getColumn(),
                                         numRows = 1,
                                         numColumns = allRange.getNumColumns())

  var columnIndexFAI = headerRange.getValues()[0].indexOf("FAI") + headerRange.getColumn()
  var columnFAI = sheetRisk.getRange(row = allRange.getRow() + 1, 
                                   column = columnIndexFAI,
                                   numRows = allRange.getNumRows() - 1,
                                   numColumns = 1)
  var columnPartLetter = sheetRisk.getRange(1,
                                          headerRange.getValues()[0].indexOf("Material") + headerRange.getColumn()
                                          ).getA1Notation().match(/([A-Z]+)/)[0]
  
  var columnFAIFormulas = []
  for (var i = 0, r = columnFAI.getRow(); i < columnFAI.getNumRows(); r ++, i ++) {
    var partcell = columnPartLetter + r
    columnFAIFormulas[i] = ['=IFERROR(INDEX(IMPORTRANGE("formulasourceformulasourceformulasource","Priority List!C:C")'+
                            ',MATCH(TRIM(' + partcell +
                            '),IMPORTRANGE("formulasourceformulasourceformulasource","Priority List!B:B"),0)),"")']
  }
  columnFAI.setValues(columnFAIFormulas)
}


function refreshFormatting() { //(sheetRisk) {
  
  
  //center columns
  let formatColumns = sheetRisk.getRange(row = allRange.getRow(), column = allRange.getColumn(),
                                           numRows = sheetRisk.getMaxRows(),
                                           numColumns = table[0].indexOf(outputfile.getRangeByName("RiskRevHeader_Action")))
  formatColumns.setHorizontalAlignment('center')

  // by name, format columns into dollars number format
  let formatCols = ["Unit Value", "Revenue"]
  for (i of formatCols) {
    formatColumns = sheetRisk.getRange(row = allRange.getRow(), column = allRange.getColumn() + table[0].indexOf(i),
                         numRows = sheetRisk.getMaxRows())
    formatColumns.setNumberFormat("$0,0")
  }
  
  formatCols = ["FAI", "HSR", "Material"]
  for (i of formatCols) {
    formatColumns = sheetRisk.getRange(row = allRange.getRow(), column = allRange.getColumn() + table[0].indexOf(i),
                         numRows = sheetRisk.getMaxRows())
    formatColumns.setNumberFormat("0")
  }

  // formatColumns = sheetRisk.getRange(row = allRange.getRow(), column = allRange.getColumn(),
  //                                          numRows = sheetRisk.getMaxRows(),
  //                                          numColumns = table[0].indexOf("Action"))

  //
  // Setting conditional formatting rules:
  // https://developers.google.com/apps-script/reference/spreadsheet/conditional-format-rule-builder
  // var ranges = [ranges] // define according to the position of the appropriate header titles
  // var rule = newConditionalFormatRule()
               //.whenCondition // not really sure how this works yet
               //.setFormat
               //.setRanges(ranges)
               //.build()
  // var rules = sheet.getConditionalFormatRules()
  // rules.push(rule)
  // sheet.setConditionalFormatRules(rules)

}
