/**
 * web_viewer.gs
 * ระบบแสดงผลตารางข่าวผ่านหน้าเว็บ (Web Viewer) 
 * สำหรับป้องกันการก๊อปปี้ข้อมูลและซ่อนไฟล์ Sheet ต้นฉบับ
 */

function doGet(e) {
  const ssId = e.parameter.id;
  if (!ssId) {
    return HtmlService.createHtmlOutput("<h1>ไม่พบข้อมูลตารางข่าว</h1><p>กรุณาคลิกลิงก์ที่ถูกต้อง</p>");
  }

  try {
    const ss = SpreadsheetApp.openById(ssId);
    
    // --- จุดปรับปรุง: ตรวจสอบว่ามีไฟล์ HTML สำเร็จรูป (Instant Load) หรือไม่ ---
    const parentFolder = DriveApp.getFileById(ssId).getParents().next();
    const htmlFileName = ss.getName() + ".html";
    const htmlFiles = parentFolder.getFilesByName(htmlFileName);
    
    if (htmlFiles.hasNext()) {
      const htmlContent = htmlFiles.next().getBlob().getDataAsString();
      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle("พชร ข่าวขายทอดตลาด เจ้าเดียวต้นฉบับ")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // --- หากไม่มีไฟล์ HTML (Fallback) ให้ทำงานแบบเดิม ---
    const sheet = ss.getSheets()[0];
    const range = sheet.getDataRange();
    
    // ดึงค่าทั้งหมดแบบ Batch (เรียกใช้ API ครั้งเดียว)
    const data = range.getValues();
    const backgrounds = range.getBackgrounds();
    const fontColors = range.getFontColors();
    const fontWeights = range.getFontWeights();
    const fontSizes = range.getFontSizes();
    const richTextValues = range.getRichTextValues();
    
    const rowCount = data.length;
    const rowHeights = [];
    
    // ลดความช้า: แทนที่จะดึงความสูงทุกแถว ให้ใช้ค่าคงที่สำหรับแถวเนื้อหา
    // แถวที่ 1 (Title) = 105, แถวที่ 2 (Header) = 45, แถววันที่ = 50, เนื้อหา = 38
    for (let i = 0; i < rowCount; i++) {
        const bg = backgrounds[i][0];
        if (i === 0) rowHeights.push(105); // Title
        else if (i === 1) rowHeights.push(45); // Header
        else if (bg === '#eeeeee') rowHeights.push(50); // Date Row
        else rowHeights.push(38); // Standard Content Row
    }

    const template = HtmlService.createTemplateFromFile('viewer_html');
    template.data = data;
    template.backgrounds = backgrounds;
    template.fontColors = fontColors;
    template.fontWeights = fontWeights;
    template.fontSizes = fontSizes;
    template.richTexts = richTextValues;
    template.rowHeights = rowHeights;
    template.colWidths = [70, 850];
    template.title = ss.getName();

    return template.evaluate()
      .setTitle("พชร ข่าวขายทอดตลาด เจ้าเดียวต้นฉบับ")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<h1>เกิดข้อผิดพลาด</h1><p>" + err.toString() + "</p>");
  }
}
