const SHOP_NAME = "";
const ACCESS_TOKEN = "";

function getLastWeekSales() {
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 14);
  const formattedDate = oneWeekAgo.toISOString();
  
  const apiVersion = '2024-01';
  const baseUrl = `https://${SHOP_NAME}.myshopify.com/admin/api/${apiVersion}`;
  
  const query = `created_at_min=${formattedDate}&status=any&fields=line_items,created_at,total_price`;
  
  const options = {
    method: 'GET',
    headers: {
      'X-Shopify-Access-Token': ACCESS_TOKEN,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(`${baseUrl}/orders.json?${query}`, options);
    const orders = JSON.parse(response.getContentText()).orders;
    
    const productSummary = {};
    
    orders.forEach(order => {
      order.line_items.forEach(item => {
        const productKey = item.variant_title ? `${item.title} (${item.variant_title})` : item.title;
        
        if (!productSummary[productKey]) {
          productSummary[productKey] = {
            productName: item.title,
            variant: item.variant_title || '-',
            quantity: 0,
            totalPrice: 0,
            price: parseFloat(item.price)
          };
        }
        
        productSummary[productKey].quantity += item.quantity;
        productSummary[productKey].totalPrice += item.quantity * parseFloat(item.price);
      });
    });
    
    const salesData = Object.values(productSummary);
    
    salesData.sort((a, b) => b.quantity - a.quantity);
    
    outputToSpreadsheet(salesData);
    
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error);
  }
}

function outputToSpreadsheet(salesData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品別販売集計') || ss.insertSheet('商品別販売集計');
  
  
  sheet.clear();
  
  const headers = ['商品名', 'バリアント', '販売数', '単価', '合計金額'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const outputData = salesData.map(item => [
    item.productName,
    item.variant,
    item.quantity,
    item.price,
    item.totalPrice
  ]);
  
  if (outputData.length > 0) {
    sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  if (outputData.length > 0) {
    sheet.getRange(2, 3, outputData.length, 1).setNumberFormat('#,##0');
    sheet.getRange(2, 4, outputData.length, 2).setNumberFormat('#,##0.00');
  }
  
  const lastRow = outputData.length + 2;
  sheet.getRange(lastRow, 1).setValue('合計');
  sheet.getRange(lastRow, 3).setFormula(`=SUM(C2:C${lastRow-1})`);
  sheet.getRange(lastRow, 5).setFormula(`=SUM(E2:E${lastRow-1})`);
  sheet.getRange(lastRow, 1, 1, headers.length).setFontWeight('bold');
}