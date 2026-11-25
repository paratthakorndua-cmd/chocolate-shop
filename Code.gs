const SHEET_NAME = "Orders";

// สร้างหัวตารางอัตโนมัติหากยังไม่มี (รันฟังก์ชันนี้ 1 ครั้งก่อน Deploy)
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  const headers = [
    "OrderID",
    "Timestamp",
    "CustomerName",
    "Phone",
    "Address",
    "SKU1_Qty",
    "SKU2_Qty",
    "TotalPrice",
    "Status"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4e342e").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
}

// รับคำสั่งซื้อใหม่จากหน้า Landing Page (POST)
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }
    
    const data = JSON.parse(e.postData.contents);
    const orderId = "ORD-" + Math.floor(100000 + Math.random() * 900000);
    const timestamp = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
    
    sheet.appendRow([
      orderId,
      timestamp,
      data.name,
      "'" + data.phone,
      data.address,
      Number(data.sku1) || 0,
      Number(data.sku2) || 0,
      Number(data.total) || 0,
      "รอหยิบสินค้า"
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", orderId: orderId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ดึงข้อมูลออเดอร์ทั้งหมด และอัปเดตสถานะการแพ็ค (GET)
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  
  const action = e && e.parameter ? e.parameter.action : null;
  
  // อัปเดตสถานะ (เช่น เปลี่ยนเป็น "แพ็คแล้ว / พร้อมส่ง")
  if (action === "updateStatus") {
    const orderId = e.parameter.orderId;
    const newStatus = e.parameter.status;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === orderId) {
        sheet.getRange(i + 1, 9).setValue(newStatus);
        return ContentService.createTextOutput(JSON.stringify({ status: "updated", orderId: orderId }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "not_found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // ดึงรายการออเดอร์ทั้งหมดไปแสดงหน้า Admin
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = rows[0];
  const orders = [];
  
  for (let i = 1; i < rows.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = rows[i][j];
    }
    orders.push(obj);
  }
  
  return ContentService.createTextOutput(JSON.stringify(orders.reverse()))
    .setMimeType(ContentService.MimeType.JSON);
}
