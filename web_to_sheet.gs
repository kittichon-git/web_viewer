/**
 * web_to_sheet.gs (V13.5 - Zero-Regex Space-Only Split)
 * - แยกคำด้วย "ช่องว่าง" (Space) อย่างเดียวเท่านั้นตามที่ผู้ใช้สั่ง
 * - จุด (.) จะถูกมองว่าเป็นส่วนหนึ่งของชื่อหน่วยงาน/รายการ ไม่ถูกแยก
 * - กฎเหล็ก: 1.วันที่(MMDD) 2.จังหวัด 3.หน่วยงาน 4+.รายการ
 */

const SECRET_KEY = "AUCTION_INTERNAL_SECRET_999";

function doGet(e) {
  const key = e.parameter.key;
  if (key !== SECRET_KEY) {
    return createJsonResponse({status: "error", message: "Unauthorized"});
  }

  const allFolderId = e.parameter.folder_all;
  const newFolderId = e.parameter.folder_new;
  const urgentFolderId = e.parameter.folder_urgent;
  const reportDateStr = e.parameter.report_date;

  let reportDate = new Date();
  if (reportDateStr) {
    const p = reportDateStr.split('-');
    reportDate = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  }
  reportDate.setHours(0, 0, 0, 0);

  const maxDate = new Date(reportDate);
  maxDate.setDate(reportDate.getDate() + 90);

  try {
    const newFiles = newFolderId ? getProcessedFiles(newFolderId, reportDate, maxDate, false) : [];
    const allFiles = allFolderId ? getProcessedFiles(allFolderId, reportDate, maxDate, true) : [];
    const urgentFiles = urgentFolderId ? getProcessedFiles(urgentFolderId, reportDate, maxDate, false) : [];

    let sheetUrl = "";
    if (newFolderId) {
      sheetUrl = createPhysicalSheet(newFiles, allFiles, reportDate, extractId(newFolderId));
    }

    const cleanForJson = (list) => {
      return list.map(f => {
        return {
          date_thai: f.date_thai || "",
          province: f.province || "",
          agency: f.agency || "",
          items: f.items || "",
          display_text: formatDisplayLine(f),
          webViewLink: f.webViewLink || "",
          sort_key: f.sort_key || ""
        };
      });
    };

    const data = {
      status: "success",
      new_files: cleanForJson(newFiles),
      all_files: cleanForJson(allFiles),
      urgent_files: cleanForJson(urgentFiles),
      sheet_url: sheetUrl
    };
    
    return createJsonResponse(data);
  } catch (err) {
    return createJsonResponse({status: "error", message: err.toString()});
  }
}

function createPhysicalSheet(newFiles, allFiles, reportDate, targetFolderId) {
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "ddMMyyyyHHmm");
  const ss = SpreadsheetApp.create("web to sheet " + timestamp);
  const sheet = ss.getActiveSheet();
  
  try {
    const file = DriveApp.getFileById(ss.getId());
    if (targetFolderId) {
      DriveApp.getFolderById(targetFolderId).addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }
  } catch (e) {}

  let row = 1;
  sheet.getRange(row, 1, 1, 2).merge().setValue("ข่าวมาใหม่ ประจำวันที่ " + formatThaiDate(reportDate)).setBackground("#FB8C00").setFontColor("#FFFFFF").setFontWeight("bold");
  row++;
  sheet.getRange(row, 1, 1, 2).setValues([["ที่", "รายละเอียดการประมูล"]]).setBackground("#EEEEEE");
  row++;

  newFiles.forEach((f, i) => {
    sheet.getRange(row, 1).setValue(i + 1);
    const text = formatDisplayLine(f);
    const richText = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(f.webViewLink).build();
    sheet.getRange(row, 2).setRichTextValue(richText);
    row++;
  });

  row += 2;
  sheet.getRange(row, 1, 1, 2).merge().setValue("ข่าวล่วงหน้า 3 เดือน").setBackground("#03A9F4").setFontColor("#FFFFFF").setFontWeight("bold");
  row++;

  allFiles.forEach((f, i) => {
    sheet.getRange(row, 1).setValue(i + 1);
    const text = formatDisplayLine(f);
    const richText = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(f.webViewLink).build();
    sheet.getRange(row, 2).setRichTextValue(richText);
    row++;
  });

  sheet.setColumnWidth(1, 45);
  sheet.setColumnWidth(2, 800);
  return ss.getUrl();
}

function formatDisplayLine(f) {
  let parts = [];
  if (f.province) parts.push(f.province);
  if (f.agency) parts.push(f.agency);
  if (f.items) parts.push(f.items);
  return parts.join("   ");
}

function getProcessedFiles(folderId, startDate, endDate, useFilter) {
  const results = [];
  const rawFiles = getFilesRecursive(extractId(folderId));
  
  rawFiles.forEach(f => {
    const parsed = parseFileName(f.name);
    if (parsed) {
      if (!useFilter || (parsed.dateObj >= startDate && parsed.dateObj <= endDate)) {
        parsed.webViewLink = f.webViewLink;
        results.push(parsed);
      }
    }
  });

  return results.sort((a, b) => {
    if (a.dateObj.getTime() !== b.dateObj.getTime()) return a.dateObj - b.dateObj;
    return a.province.localeCompare(b.province, 'th');
  });
}

function getFilesRecursive(folderId) {
  if (!folderId) return [];
  try {
    const folder = DriveApp.getFolderById(folderId);
    let results = [];
    const files = folder.getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      const file = files.next();
      results.push({ name: file.getName(), webViewLink: file.getUrl() });
    }
    const subs = folder.getFolders();
    while (subs.hasNext()) {
      results = results.concat(getFilesRecursive(subs.next().getId()));
    }
    return results;
  } catch (e) { return []; }
}

function parseFileName(fileName) {
  const nameClean = fileName.replace(/\.pdf$/i, '').trim();
  // แยกด้วยช่องว่างอย่างเดียวเท่านั้น (Space) ตัดพวก _, - ออกให้ตามสั่ง
  const parts = nameClean.split(' ').filter(p => p.trim());
  if (parts.length < 2) return null;

  const dateRaw = parts[0];
  let province = parts[1] || "";
  let agency = "";
  let items = "";

  if (parts.length >= 4) {
    agency = parts[2];
    items = parts.slice(3).join(' ');
  } else if (parts.length === 3) {
    // ถ้ามี 3 คำ: วันที่ จังหวัด รายการ (หน่วยงานข้ามไป)
    agency = ""; 
    items = parts[2];
  } else if (parts.length === 2) {
    province = parts[1];
    items = "ไม่ระบุรายละเอียด";
  }

  try {
    if (dateRaw.length < 4) return null;
    const mm = parseInt(dateRaw.substring(0, 2)) - 1;
    const dd = parseInt(dateRaw.substring(2, 4));
    let yy = 2026; 
    if (dateRaw.length >= 8) yy = parseInt(dateRaw.substring(4)) - 543;
    else if (dateRaw.length === 6) yy = parseInt("20" + dateRaw.substring(4)) - 543;
    
    const dateObj = new Date(yy, mm, dd);
    if (isNaN(dateObj.getTime())) return null;

    return {
      date_thai: formatThaiDate(dateObj),
      province: province,
      agency: agency,
      items: items,
      dateObj: dateObj,
      sort_key: Utilities.formatDate(dateObj, "GMT+7", "yyyyMMdd")
    };
  } catch (e) { return null; }
}

function formatThaiDate(date) {
  const months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  return date.getDate() + " " + months[date.getMonth()] + " " + (date.getFullYear() + 543);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function extractId(str) {
  if (!str) return "";
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str.trim();
}

function authTrigger() { DriveApp.getRootFolder(); }
