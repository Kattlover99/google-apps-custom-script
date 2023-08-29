function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('My Custom Menu')
        .addItem('Say Hello', 'helloWorld')
        .addToUi();
}

function helloWorld() {
    Browser.msgBox("Hello World!");
}