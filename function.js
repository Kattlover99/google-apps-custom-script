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

// unhide all rows and columns in current Sheet data range
function unhideRowsColumnsActiveSheet() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    sheet.unhideRow(range);
    sheet.unhideColumn(range);
}

// unhide all rows and columns in data ranges of entire Google Sheet
function unhideRowsColumnsGlobal() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
        var range = sheet.getDataRange();
        sheet.unhideRow(range);
        sheet.unhideColumn(range);
    });
};

// set all Sheets tabs to red
function setTabColor() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
        sheet.setTabColor("ff0000");
    });
};

// remove all Sheets tabs color
function resetTabColor() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
        sheet.setTabColor(null);
    });
};

function hideAllSheetsExceptActive() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
        if (sheet.getName() != SpreadsheetApp.getActiveSheet().getName())
            sheet.hideSheet();
    });
};

function unhideAllSheets() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
        sheet.showSheet();
    });
};

// reset all filters for a data range on current Sheet
function resetFilter() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    range.getFilter().remove();
    range.createFilter();
}

function DrivingMeters(origin, destination) {
    var directions = Maps.newDirectionFinder()
        .setOrigin(origin)
        .setDestination(destination)
        .getDirections();
    return directions.routes[0].legs[0].distance.value;
}

function DrivingSeconds(origin, destination) {
    var directions = Maps.newDirectionFinder()
        .setOrigin(origin)
        .setDestination(destination)
        .getDirections();
    return directions.routes[0].legs[0].duration.value;
}

function createDocument() {
    var greeting = 'Hello world!';

    var doc = DocumentApp.create('Hello_DocumentApp');
    doc.setText(greeting);
    doc.saveAndClose();
}

function getRssFeed() {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("rss-feed-contents");
    if (cached != null) {
        return cached;
    }
    // This fetch takes 20 seconds:
    var result = UrlFetchApp.fetch("http://example.com/my-slow-rss-feed.xml");
    var contents = result.getContentText();
    cache.put("rss-feed-contents", contents, 1500); // cache for 25 minutes
    return contents;
}

function getBitcoinPrice() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Get the sheet with the name Sheet1
    var sheet = spreadsheet.getSheetByName("Sheet1");
    var header = ['Timestamp', 'High', 'Low', 'Volume', 'BidAmount', 'AskAmount'];

    // Insert headers at the top row.
    sheet.getRange("A1:F1").setValues([header]);

    var url = 'https://www.bitstamp.net/api/ticker/';

    var response = UrlFetchApp.fetch(url);

    // Proceed if no error occurred.
    if (response.getResponseCode() == 200) {

        var json = JSON.parse(response);
        var result = [];

        // Timestamp
        result.push(new Date(json.timestamp *= 1000));

        // High
        result.push(json.high);

        // Low
        result.push(json.low);

        // Volume
        result.push(json.volume);

        // Bid (highest buy order)
        result.push(json.bid);

        // Ask (lowest sell order)
        result.push(json.ask);

        // Append output to Bitcoin sheet.
        sheet.appendRow(result);

    } else {

        // Log the response to examine the error
        Logger.log(response);
    }
}

function sendEmailBitcoinPricesPdfAttachment() {
    var file = SpreadsheetApp.getActiveSpreadsheet().getAs(MimeType.PDF);

    var to = 'youremail@domain.com'; // change to yours

    GmailApp.sendEmail(to, 'Bitcoin prices', 'Attached prices in PDF',
        { attachments: [file], name: 'BitcoinPrices via AppsScript' });
}

function FIRSTDAYOFTHEMONTH(year) {
    var array = [];

    for (var m = 0; m <= 11; m++) {
        var firstDay = new Date(year, m, 1);

        var dayName = '';

        switch (firstDay.getDay()) {
            case 0: dayName = 'Sunday'; break;
            case 1: dayName = 'Monday'; break;
            case 2: dayName = 'Tuesday'; break;
            case 3: dayName = 'Wednesday'; break;
            case 4: dayName = 'Thursday'; break;
            case 5: dayName = 'Friday'; break;
            case 6: dayName = 'Saturday'; break;
        }

        array.push([(m + 1) + '/1/' + year, dayName]);
    }

    return array;
}