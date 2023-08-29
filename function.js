function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('My Custom Menu')
        .addItem('Say Hello', 'helloWorld')
        .addToUi();
}

function helloWorld() {
    Browser.msgBox("Hello World!");
}

function clearInvoice() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const invoiceNumber = sheet.getRange("B5").clearContent();
    const invoiceAmount = sheet.getRange("B8").clearContent();
    const invoiceTo = sheet.getRange("E5").clearContent();
    const invoiceFrom = sheet.getRange("E6").clearContent();
}

function distanceBetweenPoints(start_point, end_point) {
    // get the directions
    const directions = Maps.newDirectionFinder()
        .setOrigin(start_point)
        .setDestination(end_point)
        .setMode(Maps.DirectionFinder.Mode.DRIVING)
        .getDirections();

    // get the first route and return the distance
    const route = directions.routes[0];
    const distance = route.legs[0].distance.text;
    return distance;
}

// function to save data
function saveData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];
    const url = sheet.getRange('Sheet1!A1').getValue();
    const follower_count = sheet.getRange('Sheet1!B1').getValue();
    const date = sheet.getRange('Sheet1!C1').getValue();
    sheet.appendRow([url, follower_count, date]);
}

// code to insert the symbol
function insertSymbol() {
    // add symbol at the cursor position
    const cursor = DocumentApp.getActiveDocument().getCursor();
    cursor.insertText('§§');
}

function logTimeRightNow() {
    const timestamp = new Date();
    Logger.log(timestamp);
}

/** @OnlyCurrentDoc */

function FormatText() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getActiveRangeList().setFontWeight('bold')
        .setFontStyle('italic')
        .setFontColor('#ff0000')
        .setFontSize(18)
        .setFontFamily('Montserrat');
};

// convert all formulas to values in the active sheet
function formulasToValuesActiveSheet() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    range.copyValuesToRange(sheet, 1, range.getLastColumn(), 1, range.getLastRow());
};

// convert all formulas to values in every sheet of the Google Sheet
function formulasToValuesGlobal() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
        var range = sheet.getDataRange();
        range.copyValuesToRange(sheet, 1, range.getLastColumn(), 1, range.getLastRow());
    });
};

// sort sheets alphabetically
function sortSheets() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
    var sheetNames = [];
    sheets.forEach(function (sheet, i) {
        sheetNames.push(sheet.getName());
    });
    sheetNames.sort().forEach(function (sheet, i) {
        spreadsheet.getSheetByName(sheet).activate();
        spreadsheet.moveActiveSheet(i + 1);
    });
};