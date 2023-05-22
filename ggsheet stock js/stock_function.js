// This function creates a custom menu in the Google Sheet.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Inventory')
    .addItem('Show Inbound Form', 'showInboundFormSidebar')
    .addItem('Show Outbound Form', 'showOutboundFormSidebar')
    .addItem('Show Transfer Form', 'showTransferFormSidebar')
    .addToUi();
}

// This function shows a sidebar with the inbound inventory form.
function showInboundFormSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('inbound')
    .setTitle('Inbound Inventory Form')
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// This function shows a sidebar with the outbound inventory form.
function showOutboundFormSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('outbound')
    .setTitle('Outbound Inventory Form')
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// This function shows a sidebar with the warehouse transfer form.
function showTransferFormSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('transfers')
    .setTitle('Warehouse Transfer Form')
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// This function adds a new inbound inventory record to the "inbound" sheet.
function addInboundInventory(formResponse) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('inbound');
  var lastRow = sheet.getLastRow();

  var primaryKey = generatePrimaryKey(); // Generate primary key
  var timestamp = new Date();
  var date = formResponse.date;
  var skuId = formResponse.skuId;
  var productName = formResponse.productName;
  var quantity = formResponse.quantity;
  var warehouse = formResponse.warehouse;
  var imageUrls = formResponse.imageUrls;

  sheet.getRange(lastRow + 1, 1).setValue(primaryKey); // Set primary key in column A
  sheet.getRange(lastRow + 1, 2).setValue(timestamp); // Set timestamp in column B
  sheet.getRange(lastRow + 1, 3).setValue(date); // Set date in column C
  sheet.getRange(lastRow + 1, 4).setValue(skuId); // Set SKU ID in column D
  sheet.getRange(lastRow + 1, 5).setValue(productName); // Set product name in column E
  sheet.getRange(lastRow + 1, 6).setValue(quantity); // Set quantity in column F
  sheet.getRange(lastRow + 1, 7).setValue(warehouse); // Set warehouse in column G
  sheet.getRange(lastRow + 1, 8).setValue(imageUrls); // Set image URLs in column H
}

// This function adds a new outbound inventory record to the "outbound" sheet.
function addOutboundInventory(formResponse) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('outbound');
  var lastRow = sheet.getLastRow();

  var primaryKey = generatePrimaryKey(); // Generate primary key
  var timestamp = new Date();
  var date = formResponse.date;
  var skuId = formResponse.skuId;
  var productName = formResponse.productName;
  var quantity = formResponse.quantity;
  var warehouse = formResponse.warehouse;

  sheet.getRange(lastRow + 1, 1).setValue(primaryKey); // Set primary key in column A
  sheet.getRange(lastRow + 1, 2).setValue(timestamp); // Set timestamp in column B
  sheet.getRange(lastRow + 1, 3).setValue(date); // Set date in column C
  sheet.getRange(lastRow + 1, 4).setValue(skuId); // Set SKU ID in column D
  sheet.getRange(lastRow + 1, 5).setValue(productName); // Set product name in column E
  sheet.getRange(lastRow + 1, 6).setValue(quantity); // Set quantity in column F
  sheet.getRange(lastRow + 1, 7).setValue(warehouse); // Set warehouse in column G
}

// This function adds a new warehouse transfer record to the "transfers" sheet.
function addTransferRecord(formResponse) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('transfers');
  var lastRow = sheet.getLastRow();

  var primaryKey = generatePrimaryKey(); // Generate primary key
  var timestamp = new Date();
  var date = formResponse.date;
  var skuId = formResponse.skuId;
  var productName = formResponse.productName;
  var quantity = formResponse.quantity;
  var originWarehouse = formResponse.originWarehouse;
  var receivedWarehouse = formResponse.receivedWarehouse;

  sheet.getRange(lastRow + 1, 1).setValue(primaryKey); // Set primary key in column A
  sheet.getRange(lastRow + 1, 2).setValue(timestamp); // Set timestamp in column B
  sheet.getRange(lastRow + 1, 3).setValue(date); // Set date in column C
  sheet.getRange(lastRow + 1, 4).setValue(skuId); // Set SKU ID in column D
  sheet.getRange(lastRow + 1, 5).setValue(productName); // Set product name in column E
  sheet.getRange(lastRow + 1, 6).setValue(quantity); // Set quantity in column F
  sheet.getRange(lastRow + 1, 7).setValue(originWarehouse); // Set origin warehouse in column G
  sheet.getRange(lastRow + 1, 8).setValue(receivedWarehouse); // Set received warehouse in column H
}


// This function generates a primary key with uppercase letters and numbers.
function generatePrimaryKey() {
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var primaryKey = '';
  for (var i = 0; i < 5; i++) {
    var randomIndex = Math.floor(Math.random() * characters.length);
    primaryKey += characters.charAt(randomIndex);
  }
  return primaryKey;
}


