function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Table Generation')
    .addItem('Create Tables From Spreadsheet Data', 'createTablesFromSpreadsheetData')
    .addItem('Delete All Tables', 'deleteAllTables')
    .addItem('Delete Tables Between Markers', 'deleteTablesBetweenMarkers')
    .addToUi();
}

function setDefaultFontToCalibri() {
  var body = DocumentApp.getActiveDocument().getBody();
  body.editAsText().setFontFamily('Calibri')
                   .setFontSize(10)
                   .setForegroundColor('#000000');
}

function insertNiskinSummaryText(position) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraph = body.insertParagraph(position, 'Niskin Sampling Summary');
  return paragraph;
}

function createTablesFromSpreadsheetData() {
  var spreadsheetId = '1CL3uzBRJDF3CSgt3w72ouKvhj0J-h9Mu4ahvA7P_rVs';
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheetNames = ['DiveReport_GEO', 'DiveReport_BIO', 'DiveReport_WAT'];
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var searchResult = body.findText("\\$SAMPLES_TABLES\\$");
  if (!searchResult) {
      Logger.log("$SAMPLES_TABLES$ not found.");
      return;
  }
  var position = body.getChildIndex(searchResult.getElement().getParent()) + 1;
  searchResult.getElement().asText().deleteText(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive());

  var geoData = [];
  var bioData = [];
  var watData = [];

  // Extract data and store in respective arrays
  for (var j = 0; j < sheetNames.length; j++) {
    var sheet = spreadsheet.getSheetByName(sheetNames[j]);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0].toString().indexOf("BLW") !== -1) {
        continue;
      }
      if (sheetNames[j] === 'DiveReport_GEO') {
        geoData.push(row);
      } else if (sheetNames[j] === 'DiveReport_BIO') {
        bioData.push(row);
      } else {
        watData.push(row);
      }
    }
  }

  // Sort the data based on naming convention
  var combinedData = geoData.concat(bioData);
  combinedData.sort(function(a, b) {
    return extractOrder(a[0]) - extractOrder(b[0]);
  });
  watData.sort(function(a, b) {
    return extractOrder(a[0]) - extractOrder(b[0]);
  });

  // Insert tables for GEO and BIO
  for (var i = 0; i < combinedData.length; i++) {
      var row = combinedData[i];
      var table = body.insertTable(position, [[]]);
      if (/_A\d{2}B$/.test(row[0])) {
          createAssociateRows(table, row, row[0].includes('B') ? 'DiveReport_BIO' : 'DiveReport_GEO');
      } else {
          createSampleRows(table, row, row[0].includes('B') ? 'DiveReport_BIO' : 'DiveReport_GEO');
          
          // Check if the next row is an associate for the current sample
          if (i + 1 < combinedData.length && !/_A\d{2}B$/.test(combinedData[i + 1][0])) {
              table = body.insertTable(position + 1, [[]]);  // Adjusted position
              createAssociateNA(table);
              position++;  // Increment position after inserting the associate table
          }
      }
      position++;  // Increment position after inserting each table
  }

  // Insert Niskin Sampling Summary text
  insertNiskinSummaryText(position);
  position++;

  // Insert tables for WAT
  for (var i = 0; i < watData.length; i++) {
    var row = watData[i];
    var table = body.insertTable(position, [[]]);
    createSampleRows(table, row, 'DiveReport_WAT');
    position++;
  }
}

function extractOrder(sampleId) {
  var match = sampleId.match(/_(\d{2}[B|G|W])/);
  if (match) {
    var order = match[1];
    return parseInt(order);
  }
  return 999;  // Default high value for unmatched items
}

function createSampleRows(table, row, sheetName) {
  var headerCell;
  if (sheetName === 'DiveReport_GEO' || sheetName === 'DiveReport_BIO') {
    var headers = ['Sample ID', 'Date (UTC)', 'Time (UTC)', 'Depth (m)', 'Latitude (decimal degrees)', 'Longitude (decimal degrees)', 'Temp. (°C)', 'Field ID(s)', 'Comments'];
    for (var i = 0; i < headers.length; i++) {
      headerCell = createRow(table, headers[i], row[i]);
      headerCell.setBackgroundColor('#00008B');
      headerCell.setForegroundColor('#FFFFFF');
    }
  } else if (sheetName === 'DiveReport_WAT') {
    var headers = ['Sample ID', 'Date (UTC)', 'Time (UTC)', 'Depth (m)', 'Latitude (decimal degrees)', 'Longitude (decimal degrees)', 'Bottle Number', 'Temperature', 'Dissolved Oxygen (mg/L)', 'Treatment'];
    for (var i = 0; i < headers.length; i++) {
      // Update here for the 'Dissolved Oxygen' value to be fetched from column M (indexed from 0)
      if (i === 8) { // 8 corresponds to the 'Dissolved Oxygen (mg/L)' header
        headerCell = createRow(table, headers[i], row[12]); // 12 corresponds to the M column
      } else if (i === 9) {
        headerCell = createRow(table, headers[i], row[16]); // As per your original code for Treatment.
      } else {
        headerCell = createRow(table, headers[i], row[i]);
      }
      headerCell.setBackgroundColor('#00008B');
      headerCell.setForegroundColor('#FFFFFF');
    }
  }
}

function createAssociateRows(table, row, sheetName) {
  var headerCell;
  headerCell = createRow(table, 'Associates Sample ID:', row[0]);
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
  headerCell = createRow(table, 'Field Identification:', row[7]);
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
  
  if (sheetName === 'DiveReport_GEO') {
    headerCell = createRow(table, 'Count:', row[11]);
  } else if (sheetName === 'DiveReport_BIO' || sheetName === 'DiveReport_WAT') {
    headerCell = createRow(table, 'Count:', row[12]);
  }
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
}

function createAssociateNA(table) {
  var headerCell;
  headerCell = createRow(table, 'Associates Sample ID:', 'N/A');
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
  headerCell = createRow(table, 'Field Identification:', 'N/A');
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
  headerCell = createRow(table, 'Count:', 'N/A');
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
}

function createRow(table, header, cellValue) {
  var row = table.appendTableRow();
  var headerCell = row.appendTableCell(header);
  var valueCell = row.appendTableCell(cellValue !== undefined ? cellValue.toString() : '');
  
  headerCell.setBackgroundColor('#00008B');
  headerCell.setForegroundColor('#FFFFFF');
  
  return headerCell;
}

function deleteTablesBetweenMarkers() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  var startPos, endPos;
  
  var foundElements = body.findText("Samples Collected");
  while (foundElements) {
    if (foundElements.getElement().asText().getFontSize(foundElements.getStartOffset()) === 18) {
      startPos = body.getChildIndex(foundElements.getElement().getParent());
      break;
    }
    foundElements = body.findText("Samples Collected", foundElements);
  }

  var endElements = body.findText("Scientists Involved");
  while (endElements) {
    if (endElements.getElement().asText().getFontSize(endElements.getStartOffset()) === 18) {
      endPos = body.getChildIndex(endElements.getElement().getParent());
      break;
    }
    endElements = body.findText("Scientists Involved", endElements);
  }

  if (startPos !== undefined && endPos !== undefined) {
    var elementsToDelete = [];
    for (var i = startPos + 1; i < endPos; i++) {
      var element = body.getChild(i);
      if (element.getType() === DocumentApp.ElementType.TABLE) {
        elementsToDelete.push(element);
      }
    }

    for (var i = 0; i < elementsToDelete.length; i++) {
      body.removeChild(elementsToDelete[i]);
    }
  }
}
