function calculateAvailableStock() {
    var inboundSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('inbound');
    var outboundSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('outbound');
    var transferSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('transfers');
    var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Available Stock');
  
    var inboundData = inboundSheet.getDataRange().getValues();
    var outboundData = outboundSheet.getDataRange().getValues();
    var transferData = transferSheet.getDataRange().getValues();
  
    var stockData = {};
  
    // Calculate inbound stock
    for (var i = 1; i < inboundData.length; i++) {
      var skuId = inboundData[i][3];
      var productName = inboundData[i][4];
      var warehouse = inboundData[i][6];
      var quantity = inboundData[i][5];
  
      if (!stockData[skuId]) {
        stockData[skuId] = {};
      }
  
      if (!stockData[skuId][warehouse]) {
        stockData[skuId][warehouse] = { quantity: quantity, productName: productName };
      } else {
        stockData[skuId][warehouse].quantity += quantity;
      }
    }
  
    // Subtract outbound stock
    for (var i = 1; i < outboundData.length; i++) {
      var skuId = outboundData[i][3];
      var warehouse = outboundData[i][6];
      var quantity = outboundData[i][5];
  
      if (stockData[skuId] && stockData[skuId][warehouse]) {
        stockData[skuId][warehouse].quantity -= quantity;
      } else {
        if (!stockData[skuId]) {
          stockData[skuId] = {};
        }
        stockData[skuId][warehouse] = { quantity: -quantity, productName: '' };
      }
    }
  
    // Add transfer stock
    for (var i = 1; i < transferData.length; i++) {
      var skuId = transferData[i][3];
      var productName = transferData[i][4];
      var originWarehouse = transferData[i][6];
      var transferToWarehouse = transferData[i][7];
      var quantity = transferData[i][5];
  
      if (stockData[skuId] && stockData[skuId][originWarehouse]) {
        stockData[skuId][originWarehouse].quantity -= quantity;
      }
  
      if (!stockData[skuId]) {
        stockData[skuId] = {};
      }
  
      if (!stockData[skuId][transferToWarehouse]) {
        stockData[skuId][transferToWarehouse] = { quantity: quantity, productName: productName };
      } else {
        stockData[skuId][transferToWarehouse].quantity += quantity;
      }
    }
  
    // Write results to the result sheet
    var result = [['SKU ID', 'Product Name', 'Warehouse', 'Available Stock']];
    for (var skuId in stockData) {
      for (var warehouse in stockData[skuId]) {
        var availableStock = stockData[skuId][warehouse].quantity;
        var productName = stockData[skuId][warehouse].productName;
        result.push([skuId, productName, warehouse, availableStock]);
      }
    }
  
    resultSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  }
  
  