const SHEET_ID = '1XMMXCloDd9-bFXzpLmcGli1rZYU5pL6g2wrxxluskLg';

// ฟังก์ชันสำหรับเสิร์ฟหน้าเว็บ และส่งข้อมูลผ่าน JSONP (GET Request)
function doGet(e) {
  // 1. ตรวจสอบว่าเป็นการขอข้อมูลรายชื่อยา (JSONP) หรือไม่
  if (e.parameter.action === 'getDrugs') {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Drugs');
    let drugs = [];
    if (sheet) {
      // ดึงข้อมูลรายชื่อยาจาก คอลัมน์ A (เริ่มแถวที่ 2)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        drugs = data.map(r => r[0]).filter(String); // กรองค่าว่างออก
      }
    }
    
    // ส่งกลับเป็น JSONP Callback
    const callback = e.parameter.callback || 'callback';
    const jsonString = JSON.stringify(drugs);
    return ContentService.createTextOutput(callback + '(' + jsonString + ')')
                         .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  // 2. ถ้าไม่มีพารามิเตอร์ ให้แสดงผลหน้าเว็บ index.html (หากใช้เป็น Web App)
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ADR Timeline Assessment')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันสำหรับรับข้อมูลและบันทึกลง Sheet (POST Request)
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data');
    if (!sheet) {
      throw new Error("ไม่พบแผ่นงานชื่อ 'Data' ใน Google Sheet");
    }
    
    // แปลงข้อมูลที่ส่งมาให้อยู่ในรูปแบบ JSON
    const data = JSON.parse(e.postData.contents);
    
    // บันทึกข้อมูลลงแถวใหม่
    sheet.appendRow([
      new Date(),       // Timestamp
      data.ptName,      // ชื่อผู้ป่วย
      data.ptHN,        // HN
      data.adrDate,     // วันที่เกิด ADR
      data.drugName,    // ชื่อยา
      data.startDate,   // วันที่เริ่มยา
      data.endDate,     // วันที่หยุดยา
      data.score,       // คะแนน Naranjo
      data.resultText   // ผลการประเมิน
    ]);
    
    // ส่งสถานะกลับไปให้หน้าเว็บ
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
