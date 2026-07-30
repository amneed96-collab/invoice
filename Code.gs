const SHEETS = {
  Admin: ["name","mobile","address","password"],
  Products: ["id","name","category","company","purchasePrice","salesPrice","stock"],
  Customers: ["id","name","mobile","shopName","address"],
  Invoices: ["id","no","date","custId","shopName","mobile","address","itemsJSON","subtotal","discount","total","paid","due","status"],
  Purchases: ["id","date","productId","name","category","company","purchasePrice","salesPrice","qty"],
  Expenses: ["id","title","amount","date","note"]
};

function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,SHEETS[name].length).setValues([SHEETS[name]]);
      sh.setFrozenRows(1);
    }
  });
  const def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
}

function doGet(e) {
  ensureSheets();
  const action = e.parameter.action;
  const sheet = e.parameter.sheet;
  if (action === "getAll") return json(getAllData());
  if (action === "get") return json(getSheetData(sheet));
  return json({error:"unknown action"});
}

function doPost(e) {
  ensureSheets();
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const sheet = body.sheet;
  if (action === "save") {
    saveSheetData(sheet, body.rows);
    return json({success:true});
  }
  if (action === "saveAdmin") {
    saveAdmin(body.data);
    return json({success:true});
  }
  return json({error:"unknown action"});
}

function getAllData() {
  const result = {};
  Object.keys(SHEETS).forEach(name => {
    if (name === "Admin") result.admin = getAdmin();
    else result[name.toLowerCase()] = getSheetData(name);
  });
  return result;
}

function getSheetData(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h,i) => {
      if (name === "Invoices" && h === "itemsJSON") {
        try { obj.items = JSON.parse(row[i] || "[]"); } catch(err) { obj.items = []; }
      } else {
        obj[h] = row[i];
      }
    });
    return obj;
  });
}

function saveSheetData(name, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  const headers = SHEETS[name];
  sh.getRange(2,1,Math.max(sh.getMaxRows()-1,1),headers.length).clearContent();
  if (rows.length === 0) return;
  const values = rows.map(r => headers.map(h => {
    if (name === "Invoices" && h === "itemsJSON") return JSON.stringify(r.items || []);
    return r[h] !== undefined ? r[h] : "";
  }));
  sh.getRange(2,1,values.length,headers.length).setValues(values);
}

function getAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Admin");
  if (sh.getLastRow() < 2) {
    sh.getRange(2,1,1,4).setValues([["Admin","","","1234"]]);
  }
  const headers = sh.getRange(1,1,1,4).getValues()[0];
  const row = sh.getRange(2,1,1,4).getValues()[0];
  const obj = {};
  headers.forEach((h,i) => obj[h] = row[i]);
  return obj;
}

function saveAdmin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Admin");
  const headers = SHEETS.Admin;
  if (sh.getLastRow() < 2) sh.getRange(2,1,1,4).setValues([["","","",""]]);
  sh.getRange(2,1,1,4).setValues([headers.map(h => data[h] !== undefined ? data[h] : "")]);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
