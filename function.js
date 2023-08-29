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