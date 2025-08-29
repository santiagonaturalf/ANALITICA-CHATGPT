// Code for Google Apps Script logic
function mainLogic() {
    // Main logic here
    Logger.log('Hello, world!');
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Run Script', 'mainLogic')
        .addToUi();
}