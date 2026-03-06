/**
 * web_to_sheet.gs (V10.3 - Simplified Urgent Mode)
 * Backend สำหรับรับค่าจากหน้าเว็บเพื่อสร้างตารางข่าว
 * 
 * [วิธีตั้งค่า]: 
 * 1. นำ Web App URL ที่ได้จากขั้นตอน Deploy ไปใส่ในบรรทัดที่ 40 (ตัวแปร webViewerUrl)
 * 2. หากต้องการทดสอบรันใน Script Editor ให้ใช้ฟังก์ชัน testManualRun ด้านล่างนี้
 */
/**
 * ฟังก์ชันสำหรับกระตุ้นให้ระบบถามหาการอนุญาต (Authorize)
 * *** ให้เลือกฟังก์ชันนี้แล้วกดปุ่ม Run (เรียกใช้) ***
 */
function authTrigger() {
  // บังคับให้ระบบขอสิทธิ์เข้าถึง Drive แบบเต็ม (รวมถึงการย้ายและลบไฟล์)
  const testFile = DriveApp.createFile("Temporary Auth Test", "test");
  testFile.setTrashed(true); // ลบทิ้งทันทีเพื่อตรวจสอบสิทธิ์การเขียน/ลบ
  
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  
  console.log("✅ อนุญาตสิทธิ์ (Permissions) สำเร็จแล้ว!");
  console.log("ตอนนี้คุณสามารถเปลี่ยนเป็นฟังก์ชัน testManualRun แล้วกด Run ได้เลยครับ");
}

function testManualRun() {
  // *** สำคัญ: ใส่เฉพาะรหัส ID (ตัวเลขตัวอักษรยาวๆ) ไม่ใช่ลิงก์ทั้งอัน ***
  // ตัวอย่างรหัส ID: 14hb0TQyuy8RzFPiv1gnoPXlgwzpxirFC
  
  const allFolderId = "14hb0TQyuy8RzFPiv1gnoPXlgwzpxirFC"; 
  const newFolderId = "10R0upMpSIniA8Yc9cw7NDGFYUWzw8QSX";
  const reportDate = "2026-03-06"; 
  
  // ฟังก์ชันช่วยตัดเอาเฉพาะ ID หากเผลอใส่มาเป็นลิงก์
  const extractId = (str) => {
    const match = str.match(/[-\w]{25,}/);
    return match ? match[0] : str;
  };

  const payload = {
    mode: "normal",
    allFolderId: extractId(allFolderId),
    newFolderId: extractId(newFolderId),
    reportDate: reportDate
  };
  
  // จำลองการเรียกใช้ doPost
  const e = {
    postData: {
      contents: JSON.stringify(payload)
    }
  };
  
  const result = doPost(e);
  console.log("ผลลัพธ์การรัน: " + result.getContent());
}

function doPost(e) {
  try {
    const contents = e.postData.contents;
    console.log("Raw received data: " + contents);
    
    const params = JSON.parse(contents);
    const mode = params.mode || 'normal'; 
    const allFolderId = params.allFolderId;
    const newFolderId = params.newFolderId;
    const urgentFolderId = params.urgentFolderId;
    const reportDateStr = params.reportDate;

    if (mode === 'normal' && !allFolderId) return createResponse({ status: "error", message: "Missing All News Folder ID" });
    
    let targetFolderId = (mode === 'urgent') ? urgentFolderId : newFolderId;
    if (!targetFolderId) return createResponse({ status: "error", message: "Missing Target Folder ID" });

    let reportDate;
    if (reportDateStr) {
      const parts = reportDateStr.split('-');
      reportDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    } else {
      reportDate = new Date();
    }
    reportDate.setHours(0, 0, 0, 0);

    // ทดสอบเรียก DriveApp ตรงนี้เพื่อกระตุ้นให้ระบบขออนุญาต (Permission)
    DriveApp.getRootFolder(); 

    const result = generateNewsReportV10_3(allFolderId, targetFolderId, reportDate, mode);
    
    // สร้างลิงก์ Web Viewer (ต้องมี ?id= ต่อท้ายด้วยครับ)
    const webViewerUrl = "https://script.google.com/macros/s/AKfycbyZCBUf7DB7GgU4GyGBp6n4C78u-RlYDu0eGv3gN6UcaEbm9outo0tKE7Ez7SPdjA7Ibw/exec?id=" + result.id;
    
    return createResponse({ 
      status: "success", 
      url: webViewerUrl, 
      sheetUrl: result.url,
      name: result.name 
    });
    
  } catch (err) {
    console.error("Critical Error: " + err.toString());
    return createResponse({ status: "error", message: "Error: " + err.toString() });
  }
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateNewsReportV10_3(allFolderId, targetFolderId, reportDate, mode) {
  const isUrgent = (mode === 'urgent');
  const typeLabel = isUrgent ? "ข่าวด่วน" : "ข่าวมาใหม่";
  const fileName = typeLabel + "_" + formatThaiDate(reportDate);
  
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const ss = SpreadsheetApp.create(fileName);
  const ssFile = DriveApp.getFileById(ss.getId());
  ssFile.moveTo(targetFolder);
  
  const sheet = ss.getActiveSheet();
  const maxDate = new Date(reportDate);
  maxDate.setDate(reportDate.getDate() + 90);

  let targetFiles = []; 
  fetchFilesFromFolder(targetFolder, targetFiles, reportDate, maxDate);

  const sorter = (a, b) => {
    if (a.dateObj.getTime() !== b.dateObj.getTime()) return a.dateObj - b.dateObj;
    if (a.province !== b.province) return a.province.localeCompare(b.province, 'th');
    return a.agency.localeCompare(b.agency, 'th');
  };

  const processedTarget = targetFiles.sort(sorter);

  sheet.clear();
  let currentRow = 1;

  // 1. ตารางข่าวหลัก (ข่าวมาใหม่ หรือ ข่าวด่วน)
  const prefix = isUrgent ? "🔥 " : "";
  const headerTarget = prefix + typeLabel + " ประจำวันที่ " + formatThaiDate(reportDate);
  currentRow = renderTableV10(sheet, headerTarget, processedTarget, currentRow, isUrgent ? "#d32f2f" : "#ff8f00", isUrgent ? "#ffebee" : "#fffde7", isUrgent); 

  // --- ข่าวด่วน: ไม่ต้องทำตารางข่าวทั้งหมด ---
  if (!isUrgent) {
    currentRow += 2; 

    let allFiles = [];
    fetchFilesFromFolder(DriveApp.getFolderById(allFolderId), allFiles, reportDate, maxDate);
    const processedAll = allFiles.sort(sorter);

    const headerAll = "📅 ข่าวล่วงหน้าทั้งหมด ประจำวันที่ " + formatThaiDate(reportDate);
    currentRow = renderTableV10(sheet, headerAll, processedAll, currentRow, "#0288d1", "#e3f2fd", false); 
  }

  applyGlobalStylingV9(sheet, currentRow);
  
  // --- ส่วนเพิ่มเติม: สร้างไฟล์ Static HTML เพื่อการโหลดที่รวดเร็ว (Instant Load) ---
  const htmlFileId = generateStaticHtmlReport(ss, targetFolder);
  
  return { url: ss.getUrl(), id: ss.getId(), name: fileName, htmlFileId: htmlFileId };
}

/**
 * ฟังก์ชันสร้างไฟล์ HTML สำเร็จรูปจากข้อมูลใน Sheet
 */
function generateStaticHtmlReport(ss, folder) {
  const sheet = ss.getSheets()[0];
  const range = sheet.getDataRange();
  const data = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontWeights = range.getFontWeights();
  const fontSizes = range.getFontSizes();
  const richTexts = range.getRichTextValues();
  
  // คำนวณความสูงแถว (ใช้ Logic เดียวกับ Web Viewer)
  const rowHeights = data.map((row, i) => {
    const bg = backgrounds[i][0];
    if (i === 0) return 105;
    if (i === 1) return 45;
    if (bg === '#eeeeee') return 50;
    return 38;
  });

  const template = HtmlService.createTemplateFromFile('viewer_html');
  template.data = data;
  template.backgrounds = backgrounds;
  template.fontColors = fontColors;
  template.fontWeights = fontWeights;
  template.fontSizes = fontSizes;
  template.richTexts = richTexts.map(row => row.map(cell => ({ linkUrl: cell.getLinkUrl() })));
  template.rowHeights = rowHeights;
  template.colWidths = [70, 850];
  template.title = ss.getName();
  template.isStatic = true; // บอก Template ว่านี่คือโหมด Static

  const htmlContent = template.evaluate().getContent();
  const htmlFileName = ss.getName() + ".html";
  
  // ลบไฟล์ HTML เก่าที่มีชื่อซ้ำในโฟลเดอร์เดียวกัน (ถ้ามี)
  const existingFiles = folder.getFilesByName(htmlFileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  
  const htmlFile = folder.createFile(htmlFileName, htmlContent, MimeType.HTML);
  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return htmlFile.getId();
}

function fetchFilesFromFolder(folder, list, startDate, endDate) {
  const files = folder.getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    const file = files.next();
    const parsedData = parseFileName(file.getName());
    if (parsedData) {
      if (parsedData.dateObj >= startDate && parsedData.dateObj <= endDate) {
        parsedData.url = file.getUrl();
        list.push(parsedData);
      }
    }
  }
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    const subFiles = sub.getFilesByType(MimeType.PDF);
    while (subFiles.hasNext()) {
      const file = subFiles.next();
      const parsedData = parseFileName(file.getName());
      if (parsedData) {
        if (parsedData.dateObj >= startDate && parsedData.dateObj <= endDate) {
          parsedData.url = file.getUrl();
          list.push(parsedData);
        }
      }
    }
  }
}

function renderTableV10(sheet, title, files, startRow, primaryColor, tableBgColor, useFireEmoji) {
  let row = startRow;
  const titleRange = sheet.getRange(row, 1, 1, 2);
  titleRange.merge().setValue("  " + title).setFontSize(24).setFontWeight("bold").setFontColor("#ffffff").setBackground(primaryColor).setVerticalAlignment("middle").setHorizontalAlignment("left"); 
  sheet.setRowHeight(row, 105); 
  row++;

  const headers = [['ที่', 'รายละเอียดการประมูล (คลิกเพื่อดูไฟล์)']];
  const headerRange = sheet.getRange(row, 1, 1, 2);
  headerRange.setValues(headers)
    .setBackground("#f5f5f5")
    .setFontColor(primaryColor)
    .setFontWeight("bold")
    .setFontSize(13)
    .setVerticalAlignment("middle");
  
  sheet.getRange(row, 1).setHorizontalAlignment("center");
  sheet.getRange(row, 2).setHorizontalAlignment("left");
  
  sheet.setRowHeight(row, 45); 
  row++;

  if (files.length === 0) {
    sheet.getRange(row, 1, 1, 2).merge().setValue("- ไม่พบข้อมูลในช่วงเวลา 90 วันนี้ -").setHorizontalAlignment("center").setFontColor("#999");
    sheet.setRowHeight(row, 45);
    return row + 1;
  }

  let lastDateStr = "";
  let itemCounter = 1;

  files.forEach(file => {
    if (lastDateStr !== file.dateThai) {
      sheet.getRange(row, 1, 1, 2).merge().setValue(" วันที่ประมูล: " + file.dateThai).setBackground("#eeeeee").setFontWeight("bold").setFontColor("#333333").setVerticalAlignment("middle");
      sheet.setRowHeight(row, 50); 
      row++;
      itemCounter = 1;
    }

    const rowRange = sheet.getRange(row, 1, 1, 2);
    rowRange.setBackground(tableBgColor); 
    sheet.setRowHeight(row, 38); 

    sheet.getRange(row, 1).setValue(itemCounter).setNumberFormat("0").setHorizontalAlignment("center").setVerticalAlignment("middle");
    const combinedText = (useFireEmoji ? "🔥 " : "") + `${file.province}   ${file.agency}   ${file.items}`;
    const richText = SpreadsheetApp.newRichTextValue().setText(combinedText).setLinkUrl(file.url).build();
    sheet.getRange(row, 2).setRichTextValue(richText).setVerticalAlignment("middle");
    rowRange.setBorder(null, null, true, null, null, null, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

    lastDateStr = file.dateThai;
    itemCounter++;
    row++;
  });
  return row;
}

function parseFileName(fileName) {
  const nameClean = fileName.replace(/\.pdf$/i, '');
  const parts = nameClean.split(' ');
  if (parts.length < 4) return null;
  const dateRaw = parts[0]; 
  const province = parts[1];
  const agency = parts[2];
  const items = parts.slice(3).join(' ');
  const month = parseInt(dateRaw.substring(0, 2)) - 1;
  const day = parseInt(dateRaw.substring(2, 4));
  const dateObj = new Date(2026, month, day); 
  return { dateObj, dateThai: formatThaiDate(dateObj), province, agency, items };
}

function formatThaiDate(date) {
  const months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const day = date.getDate(); 
  const month = months[date.getMonth()];
  const year = date.getFullYear() + 543; 
  return String(day) + " " + month + " " + year;
}

function applyGlobalStylingV9(sheet, lastRow) {
  sheet.getRange(1, 1, lastRow + 10, 2).setFontFamily('Sarabun').setFontSize(12); 
  sheet.setColumnWidth(1, 70); 
  sheet.setColumnWidth(2, 850); 
  sheet.setFrozenRows(0);
}

// --- ส่วนนี้คือไฟล์สนับสนุนที่ใช้ร่วมกัน ---

function parseFileName(fileName) {
  const nameClean = fileName.replace(/\.pdf$/i, '');
  const parts = nameClean.split(' ');
  if (parts.length < 4) return null;
  const dateRaw = parts[0]; 
  const province = parts[1];
  const agency = parts[2];
  const items = parts.slice(3).join(' ');
  const month = parseInt(dateRaw.substring(0, 2)) - 1;
  const day = parseInt(dateRaw.substring(2, 4));
  const dateObj = new Date(2026, month, day); 
  return { dateObj, dateThai: formatThaiDate(dateObj), province, agency, items };
}

function formatThaiDate(date) {
  const months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const day = date.getDate(); 
  const month = months[date.getMonth()];
  const year = date.getFullYear() + 543; 
  return String(day) + " " + month + " " + year;
}

/**
 * ฟังก์ชันสำหรับ Proxy ข้อมูลไปยัง Python (ไม่ต้องแชร์โฟลเดอร์)
 * ทดสอบเรียก: https://script.google.com/.../exec?key=AUCTION_SECRET_123
 */
function doGet(e) {
  const secretKey = "AUCTION_SECRET_123"; // รหัสผ่านความปลอดภัยที่ต้องตรงกันใน Python
  const key = e.parameter.key;
  
  if (key !== secretKey) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Unauthorized"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  const allFolderId = "14hb0TQyuy8RzFPiv1gnoPXlgwzpxirFC";
  const newFolderId = "10R0upMpSIniA8Yc9cw7NDGFYUWzw8QSX";

  try {
    const allFiles = getFilesRecursiveProxy(allFolderId);
    const newFiles = getFilesRecursiveProxy(newFolderId);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      all_files: allFiles,
      new_files: newFiles
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันช่วยดึงไฟล์แบบรวมโฟลเดอร์ย่อย (เฉพาะสำหรับ Proxy)
 */
function getFilesRecursiveProxy(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  let results = [];
  
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.PDF) {
      results.push({
        name: file.getName(),
        webViewLink: file.getUrl()
      });
    }
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    results = results.concat(getFilesRecursiveProxy(subfolders.next().getId()));
  }
  
  return results;
}
