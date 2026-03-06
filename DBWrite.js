function insertRow(sheetName, obj){

 const sheet = getSheet(sheetName);
 const headers = getHeaders(sheet);

 const row = headers.map(h => obj[h] || "");

 const last = sheet.getLastRow() + 1;

 sheet.getRange(last,1,1,row.length).setValues([row]);

}