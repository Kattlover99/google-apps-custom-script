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