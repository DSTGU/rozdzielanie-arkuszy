const master = SpreadsheetApp.getActiveSpreadsheet();
const adminSheet = master.getSheetByName("Admin");
const adminValues = adminSheet.getRange(2,1,1,5).getValues().flat();
const recieversSheet = master.getSheetByName(adminValues[0]);
const callersSheet = master.getSheetByName(adminValues[1]);

function getLastColumnInRow(sheet, rowId) {
  // Get all values in that row
  const values = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find last non-empty cell
  let lastCol = 0;
  for (let col = values.length - 1; col >= 0; col--) {
    if (values[col] !== "" && values[col] !== null) {
      lastCol = col + 1; // convert 0-based index to column number
      break;
    }
  }

  return lastCol;
}

function getLastRowInColumn(sheet, columnId) {
  const values = sheet.getRange(1, columnId, sheet.getLastRow()).getValues().flat();

  // Find last non-empty cell
  let lastRow = 0;
  for (let row = values.length - 1; row >= 0; row--) {
    if (values[row] !== "" && values[row] !== null) {
      lastRow = row + 1; // convert 0-based index to row number
      break;
    }
  }

  return lastRow;
}

function getColumnDataByHeader(sheet, headerName) {
  
  const colIndex = getColIndexByHeader(sheet, headerName);

  if (colIndex === 0) {
    throw new Error(`Column "${headerName}" not found.`);
  }

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const data = range.getValues().flat().filter(String);

  return data;
}

function getSingleDatapointByHeader(sheet, headerName) {

  const colIndex = getColIndexByHeader(sheet, headerName);

  if (colIndex === 0) {
    throw new Error(`Column "${headerName}" not found.`);
  }

  return sheet.getRange(colIndex, 2).getValue();
}

function getColIndexByHeader(sheet, headerName) {  

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName) + 1;

  if (colIndex === 0) {
    throw new Error(`Column ${headerName} not found`);
  }

  return colIndex;
}


function dividePhonesAndMails(emails, phones, phonesPerEmail = 1) {
  const availablePhones = [...phones];
  const result = {};

  emails.forEach(email => {
    const selected = [];

    for (let i = 0; i < phonesPerEmail; i++) {
      if (availablePhones.length === 0) break; // stop if we run out
      const idx = Math.floor(Math.random() * availablePhones.length);
      selected.push(availablePhones.splice(idx, 1)[0]);
    }

    result[email] = selected;
  });

  return result;
}

//GPT-izowane
function setDropdown(rangeInput, optionList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Support both string input and Range object
  const range = (typeof rangeInput === 'string')
    ? sheet.getRange(rangeInput)
    : rangeInput;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(optionList, true) // true = show dropdown UI
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
}

function getSchema() {
  const adminSheet = master.getSheetByName("Admin");
  const headers = adminSheet.getRange(5, 1, 1, getLastColumnInRow(adminSheet, 4)).getValues().flat();

  Logger.info(headers);
  
  const result = {};

  for (const idx in headers) {
    const header = headers[idx];
    result[header] = adminSheet.getRange(6, idx - -1, getLastRowInColumn(adminSheet, idx - -1) - 5).getValues().flat();
  }

  return result;
}

function generateResponseHeaders(schema) {
  const callersHeaders = callersSheet.getRange(1, 1, 1, callersSheet.getLastColumn()).getValues()[0];

  for (const key in schema) {
    if (!callersHeaders.includes(key)) {
      Logger.log("Adding entry: " + key + " to callers sheet");
      const idx = getLastColumnInRow(callersSheet, 1);
      callersSheet.getRange(1,idx+1).setValue(key);
    }
  }

}

function createAndPopulateSheets() {
  const phones = getColumnDataByHeader(recieversSheet, adminValues[2]);
  const mails = getColumnDataByHeader(callersSheet, adminValues[3]);

  const dict = dividePhonesAndMails(mails, phones, 10);
  Logger.log(dict);
  
  const callersHeaders = callersSheet.getRange(1, 1, 1, callersSheet.getLastColumn()).getValues()[0];

  if (!callersHeaders.includes(adminValues[4])) {
    Logger.log("Adding entry: " + adminValues[4] + " to callers sheet");
    const idx = getLastColumnInRow(callersSheet, 1);
    callersSheet.getRange(1,idx+1).setValue(adminValues[4]);
  }

  const colIndex = getColIndexByHeader(callersSheet, adminValues[4]);

  const schema = getSchema();
  Logger.info(schema);

  for (const idx in mails) {
    const mail = mails[idx];
    Logger.log(`Tworzenie arkusza dla ${mail}`);
    
    const child = SpreadsheetApp.create(`Obdzwanianki - ${mail}`);
    const childSheet = child.getSheets()[0];
    
    //Stworzenie kolumny z numerami
    childSheet.getRange(1,1).setValue(adminValues[2]);
    const mailPhonesRows = dict[mail].map(a => [a]);
    childSheet.getRange(2,1,dict[mail].length).setValues(mailPhonesRows);
    
    //Stworzenie pozostałych kolumn
    Object.entries(schema).forEach(([key, values], index) => {      
      childSheet.getRange(1, 2+index).setValue(key);
      setDropdown(childSheet.getRange(2, 2+index, mailPhonesRows.length, 1), values);
      childSheet.getRange(2, 2+index, mailPhonesRows.length, 1).setValue(values[values.length - 1]);
    });

    //Podziel się arkuszem
    const file = DriveApp.getFileById(child.getId());    // get the underlying file
    file.addEditor(mail); // give edit access

    //Linkowanie arkusza
    callersSheet.getRange(idx - -2, colIndex).setValue(child.getUrl());
  }

}

function getBackDataFromSheets() {
  const schema = getSchema();

  const mails = getColumnDataByHeader(callersSheet, adminValues[3]);

  const colIndex = getColIndexByHeader(callersSheet, adminValues[4]);

  generateResponseHeaders(schema);

  const links = callersSheet.getRange(2,colIndex, mails.length).getValues();

  links.forEach((link, rowIndex) => {
    Logger.log(`Odczytywanie arkusza ${link}`)
    const child = SpreadsheetApp.openByUrl(link);
    const childSheet = child.getActiveSheet();

    Object.entries(schema).forEach(([key, values]) => {      
      const data = getColumnDataByHeader(childSheet, key);
      const total = data.length;
      const correct = data.filter(t => t == values[0]).length;

      const colIndex = getColIndexByHeader(callersSheet, key);
      callersSheet.getRange(rowIndex+2, colIndex).setValue(correct/total).setNumberFormat("0.00%");

    });
  })
}



