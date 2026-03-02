// 1. ฟังก์ชันเปิดหน้าเว็บ (ห้ามเปลี่ยนชื่อ)
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Project Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ฟังก์ชันสำหรับกระตุ้นการขอสิทธิ์เข้าถึง Google Drive
function triggerAuthorization() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Authorization Successful");
}

// --- GOOGLE SHEET INTEGRATION ---
function getPriceData() {
  const SHEET_ID = '15yDGzfnX5LCfvo03kH5WaQAWSeiL5GK4Z233q8RIbWE';
  const SHEET_NAME = 'คำนวณราคา';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบชีต: ' + SHEET_NAME };

    // ดึงข้อมูลทั้งหมดตั้งแต่แถว 3 ลงมา (สมมติแถว 1-2 เป็น Header)
    const data = sheet.getRange('A3:F').getValues();

    // โครงสร้างข้อมูลที่จะส่งกลับ
    const result = [];
    let currentCategory = null;

    data.forEach(row => {
      const name = row[0]; // Col A: ชื่อรายการ
      const unit = row[2]; // Col C: หน่วย
      const price = row[3]; // Col D: ราคาต่อหน่วย

      // ข้ามแถวว่าง
      if (!name) return;

      // Logic แยกหมวดหมู่ vs สินค้า
      // ถ้า Col A มีค่า แต่ Col D (ราคา) ว่าง = เป็น "หัวข้อหมวดหมู่"
      if (name && (price === "" || price === null || price === undefined)) {
        currentCategory = {
          category: name,
          items: []
        };
        result.push(currentCategory);
      }
      // ถ้ามีราคา = เป็น "สินค้า"
      else if (price !== "" && currentCategory) {
        currentCategory.items.push({
          name: name,
          unit: unit || '',
          price: Number(price) || 0
        });
      }
    });

    return result;

  } catch (e) {
    return { error: e.message };
  }
}

// --- เพิ่มลงในไฟล์ Code.js ---

/**
 * ดึงข้อมูลรหัสวัสดุจากชีต "รหัส" แบบระบุช่องตรงๆ (แม่นยำกว่า)
 */
function getCodeMaterialData() {
  const SHEET_ID = '1pPZKZj8SvP0-_FsmEd3bT2Qb2XwyI5l5IZBH1sFoaKw';
  const SHEET_NAME = 'รหัส';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบชีต: ' + SHEET_NAME };

    // ดึงข้อมูลทั้งหมด
    const data = sheet.getDataRange().getValues();
    const result = [];

    // วนลูปเริ่มแถวที่ 2 (Index 1) เพื่อข้าม Header
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // อ้างอิงตามคอลัมน์ในชีต "รหัส" 
      // Col B (Index 1) = รหัส (New Code)
      // Col C (Index 2) = กลุ่ม
      // Col D (Index 3) = หมวด
      // Col E (Index 4) = ชื่อรายการ
      // Col G (Index 6) = หน่วย
      // Col H (Index 7) = ราคา

      const code = row[1];
      const name = row[4];

      // ถ้าไม่มีทั้งรหัสและชื่อ ให้ข้ามบรรทัดนี้
      if (!code && !name) continue;

      result.push({
        code: code ? code.toString() : '-',
        group: row[2] ? row[2].toString() : '',
        category: row[3] ? row[3].toString() : '',
        name: name ? name.toString() : '',
        unit: row[6] ? row[6].toString() : '',
        price: row[7] || 0
      });
    }

    return result;

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * บันทึกข้อมูลวัสดุใหม่
 */
function createMaterial(formData) {
  const SHEET_ID = '1pPZKZj8SvP0-_FsmEd3bT2Qb2XwyI5l5IZBH1sFoaKw';
  const SHEET_NAME = 'รหัส';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบชีต: ' + SHEET_NAME };

    // เตรียมข้อมูลลงแถว (ว่าง, Code, Group, Category, Name, ว่าง, Unit, Price)
    // Index: 0=A, 1=B, 2=C, 3=D, 4=E, 5=F, 6=G, 7=H
    const newRow = [
      '',                     // A
      formData.code,          // B
      formData.group,         // C
      formData.category,      // D
      formData.name,          // E
      '',                     // F
      formData.unit,          // G
      formData.price          // H
    ];

    sheet.appendRow(newRow);
    return { success: true };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * อัปเดตข้อมูลวัสดุเดิม (ค้นหาด้วยรหัสเดิม)
 */
function updateCodeMaterial(formData) {
  const SHEET_ID = '1pPZKZj8SvP0-_FsmEd3bT2Qb2XwyI5l5IZBH1sFoaKw';
  const SHEET_NAME = 'รหัส';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบชีต: ' + SHEET_NAME };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // ค้นหาแถวที่รหัสตรงกัน (เริ่มแถว 2)
    // ใช้ formData.originalCode หากมีการเปลี่ยนรหัส, หรือใช้ formData.code แทน
    const searchCode = formData.originalCode || formData.code;

    for (let i = 1; i < data.length; i++) {
      // เช็ค Col B (Index 1)
      if (data[i][1].toString() === searchCode.toString()) {
        rowIndex = i + 1; // 1-based index because data array is 0-based but Sheet rows are 1-based
        break;
      }
    }

    if (rowIndex === -1) return { error: 'ไม่พบรหัสสินค้า: ' + searchCode };

    // อัปเดตข้อมูลเฉพาะ Cell ที่เกี่ยวข้อง
    // B=Code, C=Group, D=Category, E=Name, G=Unit, H=Price
    // getRange(row, column) -> column index starts at 1 (A=1, B=2...)

    // Update Code (B=2)
    sheet.getRange(rowIndex, 2).setValue(formData.code);
    // Update Group (C=3)
    sheet.getRange(rowIndex, 3).setValue(formData.group);
    // Update Category (D=4)
    sheet.getRange(rowIndex, 4).setValue(formData.category);
    // Update Name (E=5)
    sheet.getRange(rowIndex, 5).setValue(formData.name);
    // Update Unit (G=7)
    sheet.getRange(rowIndex, 7).setValue(formData.unit);
    // Update Price (H=8)
    sheet.getRange(rowIndex, 8).setValue(formData.price);

    return { success: true };

  } catch (e) {
    return { error: e.message };
  }
}

// --- ดึงข้อมูลจากชีต FM-GP-03 สำหรับเมนู Tasks ---
function getTasksData() {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA FM-GP-02';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบชีต: ' + SHEET_NAME };

    const data = sheet.getDataRange().getDisplayValues();
    const result = [];

    // แก้ไข: เปลี่ยนจาก i = 1 เป็น i = 2 เพื่อข้ามแถวที่ 1 และ 2 ของสเปรดชีต (Index 0 และ 1)
    for (let i = 2; i < data.length; i++) {
      const row = data[i];

      // ดักจับเพิ่มเติม: ถ้าไม่มีข้อมูลในคอลัมน์แรก หรือค่าเป็นคำว่า "No. Order" ให้ข้ามไปเลย
      if (!row[0] || row[0] === 'No. Order') continue;

      result.push({
        no: row[0] || '-',               // A: No. Order
        date: row[1] || '-',             // B: Date Create
        product: row[2] || '-',          // C: Product Name
        customer: row[3] || '-',         // D: Customer
        qty: row[4] || '',               // E: QTY
        id: row[5] || '',                // F: ID (Internal)
        sales: row[6] || '-',            // G: Sale Name
        designer: row[7] || '-',         // H: Graphic Name
        period: row[8] || '',            // I: Production Period
        installDays: row[9] || '',       // J: Install Days
        totalTime: row[10] || '',        // K: Total Days
        status: row[11] || 'รอดำเนินการ',   // L: Status Update
        detail: row[12] || '-',          // M: Remark
        user: row[13] || '-',            // N: Creator Name
        update: row[14] || '-',          // O: User_Update
        driveLink: row[15] || '',         // P: Drive Folder Link
        startTime: row[16] || '',        // Q: Start Counting Time
        endTime: row[17] || ''           // R: End Time
      });
    }

    // เรียงให้ข้อมูลล่าสุด (ใบงานใหม่) ขึ้นก่อน
    return result.reverse();

  } catch (e) {
    return { error: e.message };
  }
}

// --- บันทึกใบงานใหม่ลงชีต FM-GP-03 ---

// ==========================================
// ฟังก์ชันช่วยเหลือ (Helper Functions)
// ==========================================

// 1. ฟังก์ชันสร้างวันที่และเวลาแบบไทย (พ.ศ.)
function getThaiDateTime() {
  const now = new Date();
  // d/M/yyyy จะได้รูปแบบ เช่น 18/2/2026
  const dateStr = Utilities.formatDate(now, "GMT+7", "d/M/yyyy");
  const timeStr = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");

  const dateParts = dateStr.split('/');
  const yearBE = parseInt(dateParts[2]) + 543; // แปลง ค.ศ. เป็น พ.ศ.

  const formattedDate = `${dateParts[0]}/${dateParts[1]}/${yearBE}`;

  return {
    fullDateTime: `${formattedDate} ${timeStr}`,           // ผลลัพธ์: 18/2/2569 14:21:00
    updateFormat: `${formattedDate}, ${timeStr}`           // ผลลัพธ์: 18/2/2569, 14:21:00
  };
}

// 2. ฟังก์ชันรันหมายเลข No. Order (รีเซ็ตทุกเดือน)
function generateOrderNumber(sheet) {
  const now = new Date();
  // รหัส No. Order ใช้ ปี ค.ศ. ตามที่ระบุ
  const yearCE = Utilities.formatDate(now, "GMT+7", "yyyy");
  const month = Utilities.formatDate(now, "GMT+7", "MM");
  const day = Utilities.formatDate(now, "GMT+7", "dd");

  const todayPrefix = `${yearCE}${month}${day}`; // เช่น 20260218
  const currentMonthPrefix = `${yearCE}${month}`; // เช่น 202602 สำหรับเช็คว่าเดือนเดียวกันไหม

  // ดึงข้อมูล No. Order ทั้งหมดในคอลัมน์ A (เริ่มแถว 3) มาตรวจสอบ
  const data = sheet.getRange("A3:A").getValues().flat().filter(String);
  let maxSeq = 0;

  for (let i = 0; i < data.length; i++) {
    const orderNo = data[i].toString();
    // ถ้าใบงานเก่าอยู่ใน "เดือนปัจจุบัน" ให้นำเลขต่อท้ายมาเทียบหาค่าสูงสุด
    if (orderNo.substring(0, 6) === currentMonthPrefix) {
      const parts = orderNo.split('-');
      if (parts.length === 2) {
        const seq = parseInt(parts[1], 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    }
  }

  // บวก 1 จากค่าสูงสุดของเดือนนี้ (ถ้าขึ้นเดือนใหม่ maxSeq จะเป็น 0 และเริ่มที่ 001)
  const nextSeq = (maxSeq + 1).toString().padStart(3, '0');
  return `${todayPrefix}-${nextSeq}`;
}


// ==========================================
// ฟังก์ชันหลักสำหรับหน้า Tasks
// ==========================================

// --- บันทึกใบงานใหม่ลงชีต FM-GP-03 ---
function createTaskRecord(formData) {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA FM-GP-02';

  // ใช้ LockService เพื่อป้องกันคนกดเซฟพร้อมกันแล้วได้เลข Order ซ้ำ
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000); // รอคิวรันโค้ดสูงสุด 10 วินาที

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // สร้างข้อมูลอัตโนมัติจากฝั่งเซิร์ฟเวอร์
    const newOrderNo = generateOrderNumber(sheet);
    const thaiTime = getThaiDateTime();
    const creatorName = formData.user || 'Unknown'; // รับชื่อ User จากหน้าเว็บ
    const p = parseFloat(formData.period) || 0;
    const iDays = parseFloat(formData.installDays) || 0;
    const totalDays = p + iDays;

    const status = formData.status || 'รอดำเนินการ';
    const startTime = (status === 'รอดำเนินการ') ? thaiTime.fullDateTime : '';

    const newRow = [
      newOrderNo,             // A: No. Order (รันออโต้)
      thaiTime.fullDateTime,  // B: Date Create (เวลาพ.ศ. ออโต้)
      formData.product,       // C: Product Name
      formData.customer,      // D: Customer
      formData.qty || '',     // E: QTY
      formData.id || '',      // F: Internal ID
      formData.sales,         // G: Sale Name
      formData.designer,      // H: Graphic Name
      formData.period || '0', // I: Production Period
      formData.installDays || '0', // J: Install Days
      totalDays,              // K: Total Duration
      status,                 // L: Status Update
      formData.detail || '',  // M: Remark
      creatorName,            // N: Creator Name
      '-',                    // O: User_Update
      '',                     // P: Drive Folder Link
      startTime               // Q: Start Counting Time (รันออโต้เฉพาะเมื่อสถานะเป็น รอดำเนินการ)
    ];

    sheet.appendRow(newRow);

    // ส่งข้อมูลที่สร้างเสร็จกลับไปให้หน้าเว็บอัปเดตตารางทันที
    return {
      success: true,
      generatedData: {
        no: newOrderNo,
        date: thaiTime.fullDateTime,
        user: creatorName,
        update: '-',
        startTime: startTime
      }
    };

  } catch (e) {
    return { error: e.message };
  } finally {
    lock.releaseLock(); // ปล่อยคิวเมื่อทำงานเสร็จ
  }
}

// --- อัปเดตข้อมูลและบันทึก User_Update ---
function updateTaskRecord(orderNo, updateData, userName) {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA FM-GP-02';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    const data = sheet.getRange("A:A").getValues();
    let rowIndex = -1;

    for (let i = 2; i < data.length; i++) {
      if (data[i][0] === orderNo) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: 'ไม่พบหมายเลขใบงาน' };

    const thaiTime = getThaiDateTime();
    const updateLogStr = `${userName}, ${thaiTime.updateFormat}`;

    // A: No. Order = 1
    // B: Date Create = 2
    // C: Product = 3
    // D: Customer = 4
    // E: QTY = 5
    // F: ID = 6
    // G: Sale Name = 7
    // H: Graphic Name = 8
    // I: Production Period = 9
    // J: Install Days = 10
    // K: Total Days = 11
    // L: Status Update = 12
    // M: Remark = 13
    // N: Creator = 14
    // O: User_Update = 15
    // P: Drive Link = 16

    if (updateData.product !== undefined) sheet.getRange(rowIndex, 3).setValue(updateData.product);
    if (updateData.customer !== undefined) sheet.getRange(rowIndex, 4).setValue(updateData.customer);
    if (updateData.qty !== undefined) sheet.getRange(rowIndex, 5).setValue(updateData.qty);
    if (updateData.id !== undefined) sheet.getRange(rowIndex, 6).setValue(updateData.id);
    if (updateData.sales !== undefined) sheet.getRange(rowIndex, 7).setValue(updateData.sales);
    if (updateData.designer !== undefined) sheet.getRange(rowIndex, 8).setValue(updateData.designer);
    if (updateData.period !== undefined) sheet.getRange(rowIndex, 9).setValue(updateData.period);
    if (updateData.installDays !== undefined) sheet.getRange(rowIndex, 10).setValue(updateData.installDays);

    // Auto-update total in K if period or installDays changed
    const currentPeriod = sheet.getRange(rowIndex, 9).getValue();
    const currentInstall = sheet.getRange(rowIndex, 10).getValue();
    sheet.getRange(rowIndex, 11).setValue((parseFloat(currentPeriod) || 0) + (parseFloat(currentInstall) || 0));

    if (updateData.status !== undefined) {
      sheet.getRange(rowIndex, 12).setValue(updateData.status);

      // บันทึกเวลาที่คอลัมน์ Q (Index 17) หากเปลี่ยนเป็น "รอดำเนินการ" และยังไม่มีข้อมูล
      if (updateData.status === 'รอดำเนินการ') {
        const currentStartTime = sheet.getRange(rowIndex, 17).getValue();
        if (!currentStartTime || currentStartTime === '-') {
          sheet.getRange(rowIndex, 17).setValue(thaiTime.fullDateTime);
        }
      }

      // บันทึกเวลาที่คอลัมน์ R (Index 18) เมื่อสถานะเป็น เสร็จสิ้น/Completed
      if (['ดำเนินการแล้วเสร็จ', 'Completed', 'เสร็จสิ้น'].includes(updateData.status)) {
        sheet.getRange(rowIndex, 18).setValue(thaiTime.fullDateTime);
      }
    }

    if (updateData.detail !== undefined) sheet.getRange(rowIndex, 13).setValue(updateData.detail);


    const existingUpdateLog = sheet.getRange(rowIndex, 15).getValue();
    const finalUpdateLog = (existingUpdateLog && existingUpdateLog !== '-')
      ? existingUpdateLog + ", " + updateLogStr
      : updateLogStr;

    sheet.getRange(rowIndex, 15).setValue(finalUpdateLog);

    return { success: true, newUpdateLog: finalUpdateLog };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

const ATTACHMENT_ROOT_FOLDER_ID = '15T7qgwYIVHTWICZB3X6fbPMuBx0pm7zC';

/**
 * สร้างโฟลเดอร์และอัปโหลดรูปภาพ
 */
function uploadTaskImages(base64Images, orderNo) {
  try {
    const rootFolder = DriveApp.getFolderById(ATTACHMENT_ROOT_FOLDER_ID);
    let orderFolder;

    // Check if folder already exists
    const folders = rootFolder.getFoldersByName(orderNo);
    if (folders.hasNext()) {
      orderFolder = folders.next();
    } else {
      orderFolder = rootFolder.createFolder(orderNo);
    }

    const imageFolder = orderFolder.createFolder('รูปภาพประกอบใบงาน');

    base64Images.forEach((base64Data, index) => {
      const parts = base64Data.split(',');
      const meta = parts[0];
      const contentType = meta.substring(5, meta.indexOf(';'));
      const bytes = Utilities.base64Decode(parts[1]);
      const ext = getExtensionFromMime(contentType);
      const filename = `attachment_${index + 1}${ext}`;
      const blob = Utilities.newBlob(bytes, contentType, filename);
      imageFolder.createFile(blob);
    });

    const folderUrl = orderFolder.getUrl();
    updateTaskFolderLink(orderNo, folderUrl);

    return { success: true, folderUrl: folderUrl };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * อัปโหลดรูปภาพงานที่เสร็จสิ้น (FM-GP-03)
 */
function uploadCompletionImages(base64Images, orderNo) {
  try {
    const rootFolder = DriveApp.getFolderById(ATTACHMENT_ROOT_FOLDER_ID);
    let orderFolder;

    // Find task folder
    const folders = rootFolder.getFoldersByName(orderNo.trim());
    if (folders.hasNext()) {
      orderFolder = folders.next();
    } else {
      orderFolder = rootFolder.createFolder(orderNo.trim());
    }

    // Find or Create completion folder
    let completionFolder;
    const subFolders = orderFolder.getFoldersByName('FM-GP-03 ใบงานอนุมัติงานผลิต');
    if (subFolders.hasNext()) {
      completionFolder = subFolders.next();
    } else {
      completionFolder = orderFolder.createFolder('FM-GP-03 ใบงานอนุมัติงานผลิต');
    }

    // Clear existing files in the specific completion subfolder before uploading new ones
    const filesInFolder = completionFolder.getFiles();
    while (filesInFolder.hasNext()) {
      filesInFolder.next().setTrashed(true);
    }

    base64Images.forEach((base64Data, index) => {
      const parts = base64Data.split(',');
      const meta = parts[0];
      const contentType = meta.substring(5, meta.indexOf(';'));
      const bytes = Utilities.base64Decode(parts[1]);
      const ext = getExtensionFromMime(contentType);
      const filename = `completion_${index + 1}${ext}`;
      const blob = Utilities.newBlob(bytes, contentType, filename);
      completionFolder.createFile(blob);
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * อัปเดตลิงก์โฟลเดอร์ใน Sheet
 */
function updateTaskFolderLink(orderNo, folderUrl) {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA FM-GP-02';

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getRange("A:A").getValues();

  for (let i = 2; i < data.length; i++) {
    if (data[i][0] === orderNo) {
      sheet.getRange(i + 1, 14).setValue(folderUrl); // คอลัมน์ N = 14
      break;
    }
  }
}

// Helper สำหรับระบุสกุลไฟล์ตาม MIME type
function getExtensionFromMime(mime) {
  const mimeMap = {
    'image/jpeg': '.jpg',
    'image/png': '.png',
    'image/gif': '.gif',
    'application/pdf': '.pdf',
    'application/zip': '.zip',
    'application/x-zip-compressed': '.zip',
    'application/x-rar-compressed': '.rar',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
    'application/msword': '.doc'
  };
  return mimeMap[mime] || '.bin';
}

/**
 * ดึงภาพใบงาน (สำหรับ Preview)
-03 จาก Google Drive เป็น Base64
 */
function getTaskCompletionImages(orderNo) {
  try {
    const rootFolder = DriveApp.getFolderById(ATTACHMENT_ROOT_FOLDER_ID);
    const orderNoStr = String(orderNo).trim();
    const folders = rootFolder.getFoldersByName(orderNoStr);
    const images = [];

    // แสกนทุกโฟลเดอร์ที่มีชื่อตรงกันเพื่อรวมรูป (เผื่อกรณีมีโฟลเดอร์ซ้ำซ้อนจากบั๊กก่อนหน้า)
    while (folders.hasNext()) {
      const orderFolder = folders.next();
      const subFolders = orderFolder.getFoldersByName('FM-GP-03 ใบงานอนุมัติงานผลิต');

      while (subFolders.hasNext()) {
        const completionFolder = subFolders.next();
        const files = completionFolder.getFiles();

        while (files.hasNext()) {
          const file = files.next();
          const blob = file.getBlob();
          const contentType = blob.getContentType();
          if (contentType.indexOf('image/') !== -1) {
            const bytes = blob.getBytes();
            const base64 = Utilities.base64Encode(bytes);
            images.push('data:' + contentType + ';base64,' + base64);
          }
        }
      }
    }

    // เรียงลำดับรูปภาพ (optional - ตามชื่อไฟล์เพื่อให้ลำดับเหมือนตอนอัปโหลด)
    return images;
  } catch (e) {
    Logger.log("Error fetching images: " + e.message);
    return [];
  }
}

// ==========================================
// ==========================================
// ดึงข้อมูลเนื้อหาในโฟลเดอร์ (ทั้งโฟลเดอร์ย่อยและไฟล์)
// ==========================================
function getDriveFolderContent(folderId) {
  try {
    const parentId = folderId || '15T7qgwYIVHTWICZB3X6fbPMuBx0pm7zC';
    const folder = DriveApp.getFolderById(parentId);

    const result = {
      success: true,
      folderName: folder.getName(),
      folders: [],
      files: []
    };

    // ดึงโฟลเดอร์ย่อย
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const sub = subfolders.next();
      result.folders.push({
        id: sub.getId(),
        name: sub.getName(),
        lastUpdated: sub.getLastUpdated().toISOString(),
        type: 'folder'
      });
    }

    // ดึงไฟล์
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      result.files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        size: file.getSize(),
        lastUpdated: file.getLastUpdated().toISOString(),
        downloadUrl: `https://drive.google.com/uc?export=download&id=${file.getId()}`,
        url: file.getUrl(),
        type: 'file'
      });
    }

    // เรียงลำดับ (ตามชื่อ)
    result.folders.sort((a, b) => b.name.localeCompare(a.name)); // โฟลเดอร์งานล่าสุดขึ้นก่อน
    result.files.sort((a, b) => a.name.localeCompare(b.name));

    return result;
  } catch (e) {
    console.error("Error fetching drive content:", e);
    return { success: false, error: e.message };
  }
}

// Keep the old ones for compatibility if needed, but point to the new one or replace them completely.
function getDriveFoldersData() { return getDriveFolderContent(); }
function getDriveFilesData(folderId) { return getDriveFolderContent(folderId); }

// ==========================================
// ระบบ เข้าสู่ระบบ (Login)
// ==========================================
function checkLogin(username, password) {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA Password'; // ต้องตั้งชื่อชีตให้ตรงกับใน Google Sheet

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) return { success: false, message: 'ไม่พบฐานข้อมูลผู้ใช้งาน (ชีต DATA Password)' };

    // ดึงข้อมูลทั้งหมดในชีตมาตรวจสอบ
    const data = sheet.getDataRange().getValues();

    // วนลูปเช็คข้อมูลทีละแถว (เริ่มที่ 1 เพื่อข้ามหัวข้อในแถวแรก)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sheetUsername = row[1] ? row[1].toString().trim() : ''; // Col B
      const sheetPassword = row[2] ? row[2].toString().trim() : ''; // Col C
      const sheetRole = row[3] ? row[3].toString().trim() : 'Staff'; // Col D
      const sheetAvatar = row[4] ? row[4].toString().trim() : null; // Col E

      // ถ้า Username และ Password ตรงกัน
      if (sheetUsername === username.trim() && sheetPassword === password.trim()) {
        return {
          success: true,
          user: {
            name: sheetUsername,
            role: sheetRole,
            avatar: sheetAvatar
          }
        };
      }
    }

    // ถ้าวนลูปจนจบแล้วหาไม่เจอ
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };

  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการเชื่อมต่อ: ' + e.message };
  }
}

// ฟังก์ชันดึงรายชื่อผู้ใช้ทั้งหมด (สำหรับ Autocomplete)
function getUserList() {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA Password';
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    // ดึงข้อมูลทั้งหมด ข้ามหัวข้อ
    const data = sheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < data.length; i++) {
      // Column B = Username (Index 1)
      const username = data[i][1] ? data[i][1].toString().trim() : '';
      if (username) {
        users.push(username);
      }
    }
    return users;
  } catch (e) {
    return [];
  }
}

// --- Material Calculation Functions (Slim Lightbox & Others) ---

const SHEET_CALC_ID = '1AkUXwcD89sLdjyfYY0bZHX-MCanYt9Plk8od_9swLUw'; // ID for Calculation Sheet

function getSheetForCalculation(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_CALC_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'ไม่พบชีต: ' + sheetName };

    const range = sheet.getDataRange();
    // ดึงค่าที่แสดงผล (มีคอมม่า จุดทศนิยม) และดึงสูตร เพื่อแยกว่าช่องไหนแก้ไขได้/ไม่ได้
    return {
      success: true,
      displayValues: range.getDisplayValues(),
      formulas: range.getFormulas()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateCellAndRecalculate(sheetName, rowIdx, colIdx, value) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_CALC_ID);
    const sheet = ss.getSheetByName(sheetName);

    // อัปเดตค่าลง Sheet (บวก 1 เพราะ Index ใน Sheet เริ่มที่ 1 แต่ใน Array เริ่มที่ 0)
    sheet.getRange(rowIdx + 1, colIdx + 1).setValue(value);

    // บังคับให้ Google Sheet คำนวณสูตรใหม่ทันที
    SpreadsheetApp.flush();

    // ดึงข้อมูลที่คำนวณเสร็จแล้วกลับไปแสดง
    return getSheetForCalculation(sheetName);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function checkSlimLightBoxFormulas() {
  const SHEET_NAME = 'คำนวณราคา Slim light box';

  try {
    const ss = SpreadsheetApp.openById(SHEET_CALC_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log('ไม่พบชีตชื่อ: ' + SHEET_NAME);
      return;
    }

    const range = sheet.getDataRange();
    const formulas = range.getFormulas();

    Logger.log('=== รายการสูตรในชีต: ' + SHEET_NAME + ' ===');
    let formulaCount = 0;

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        const formula = formulas[r][c];
        if (formula !== "") {
          formulaCount++;
          let colLetter = getColumnLetter(c + 1);
          let rowNum = r + 1;
          Logger.log(`เซลล์ ${colLetter}${rowNum} ใช้สูตร: ${formula}`);
        }
      }
    }

    Logger.log('===================================');
    Logger.log(`พบสูตรทั้งหมด: ${formulaCount} ช่อง`);

  } catch (e) {
    Logger.log('เกิดข้อผิดพลาด: ' + e.message);
  }
}

function getColumnLetter(columnNumber) {
  let temp, letter = '';
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

// --- Slim Lightbox Data Fetching ---
function getSlimLightboxRawData() {
  const SPREADSHEET_ID = '1BPk2b_KbY4cqW_Q5YTruDUgJ0zOLemadnPO56R4qSwg';
  const SHEET_NAME = 'Raw data Slim light box'; // ดึงตามชื่อชีตให้ชัวร์

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // หาชีตจากชื่อแทน GID เพื่อป้องกันผิดแท็บ
    let sheet = ss.getSheetByName(SHEET_NAME);
    let allNames = [];

    // ถ้าหาชื่อเป๊ะๆ ไม่เจอ ลองหาแบบตัดเว้นวรรค
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const sName = sheets[i].getName();
        allNames.push(sName);
        if (sName.toLowerCase().replace(/\s/g, '') === SHEET_NAME.toLowerCase().replace(/\s/g, '')) {
          sheet = sheets[i];
          break;
        }
      }
    }

    if (!sheet) {
      return { error: 'หาหน้าชีตชื่อ "' + SHEET_NAME + '" ไม่เจอครับ. หน้าที่มีตอนนี้คือ: ' + allNames.join(', ') };
    }

    // Fetch all data including headers
    let data = sheet.getDataRange().getValues();

    // Convert Dates to timestamp strings to prevent google.script.run silent failure
    data = data.map(row => row.map(cell => {
      if (cell instanceof Date) return cell.toISOString();
      return cell;
    }));

    return { success: true, data: data };

  } catch (e) {
    return { error: e.message };
  }
}

// --- ฟังก์ชันดึงฐานข้อมูลงานพิมพ์ (ดึงทีเดียว 2 ชีต) ---
function getPrintingRawData() {
  const SPREADSHEET_ID = '1BPk2b_KbY4cqW_Q5YTruDUgJ0zOLemadnPO56R4qSwg';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rawSheet = ss.getSheetByName('Raw data งานพิมพ์');
    const helperSheet = ss.getSheetByName('ช่วยคำนวนงานพิมพ์');

    if (!rawSheet || !helperSheet) {
      return { error: 'ไม่พบชีต Raw data งานพิมพ์ หรือ ช่วยคำนวนงานพิมพ์' };
    }

    // ดึงข้อมูลทั้งหมดมาเป็น Array (เปลี่ยนใช้ getRange แบบ Fix คอลัมน์ไปจนถึง AE เพื่อป้องกันข้อมูลหาย)
    return {
      success: true,
      rawData: rawSheet.getRange('A1:AE100').getValues(),
      helperData: helperSheet.getRange('A1:AE20').getValues()
    };

  } catch (e) {
    return { error: e.message };
  }
}

// --- ฟังก์ชันเช็คสูตรของหน้า คำนวนงานพิมพ์ ---
function checkPrintingFormulas() {
  const SPREADSHEET_ID = '1BPk2b_KbY4cqW_Q5YTruDUgJ0zOLemadnPO56R4qSwg';
  const SHEET_NAME = 'คำนวนงานพิมพ์';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log('ไม่พบชีตชื่อ: ' + SHEET_NAME);
      return;
    }

    const formulas = sheet.getDataRange().getFormulas();
    Logger.log('=== รายการสูตรในชีต: ' + SHEET_NAME + ' ===');

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c] !== "") {
          let colLetter = getColumnLetter(c + 1);
          Logger.log(`เซลล์ ${colLetter}${r + 1} ใช้สูตร: ${formulas[r][c]}`);
        }
      }
    }
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}

// --- ฟังก์ชันเช็ครายชื่อวัสดุทั้งหมด (สำหรับทำ Dropdown) ---
function checkPrintingMaterials() {
  const SPREADSHEET_ID = '1BPk2b_KbY4cqW_Q5YTruDUgJ0zOLemadnPO56R4qSwg';
  const SHEET_NAME = 'ช่วยคำนวนงานพิมพ์';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log('ไม่พบชีตชื่อ: ' + SHEET_NAME);
      return;
    }

    // ดึงข้อมูลแถวที่ 2 (ซึ่งเก็บชื่อวัสดุไว้) เริ่มตั้งแต่คอลัมน์ A ถึง ZZ
    const row2 = sheet.getRange("A2:ZZ2").getValues()[0];

    Logger.log('=== รายการประเภทวัสดุที่ใช้พิมพ์ (จากแถวที่ 2) ===');

    let materials = [];
    // เริ่มเช็คตั้งแต่คอลัมน์ AE (index 30) ไปเรื่อยๆ
    for (let i = 30; i < row2.length; i++) {
      if (row2[i] !== "") {
        materials.push(row2[i]);
        // เรียกใช้ฟังก์ชัน getColumnLetter ที่คุณมีอยู่แล้วเพื่อดูคอลัมน์
        let colLetter = getColumnLetter(i + 1);
        Logger.log(`คอลัมน์ ${colLetter}: ${row2[i]}`);
      }
    }

    Logger.log('-----------------------------------');
    Logger.log('สรุปวัสดุทั้งหมดสำหรับนำไปทำ Dropdown:');
    Logger.log(JSON.stringify(materials));

  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}

// ==========================================
// ฟังก์ชันดึงรายชื่อ Sales (สำหรับ Tasks Dropdown)
// ==========================================
function getSalesUserList() {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA Password';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    // ดึงข้อมูลทั้งหมด (A:Username, B:Password, C:Position)
    const data = sheet.getDataRange().getValues();
    const salesUsers = [];

    // เริ่มที่ i = 1 (ข้ามหัวข้อ)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const username = row[1] ? String(row[1]).trim() : ''; // Col B
      const position = row[3] ? String(row[3]).trim().toLowerCase() : ''; // Col D = Position

      // กรองเฉพาะตำแหน่งที่มีคำว่า "sale" หรือ "เซลล์"
      if (username && (position.includes('sale') || position.includes('เซลล์'))) {
        salesUsers.push(username);
      }
    }

    // เรียงตามตัวอักษร
    return salesUsers.sort();

  } catch (e) {
    Logger.log("Error getting sales users: " + e.message);
    return [];
  }
}

// ฟังก์ชันดึงรายชื่อ Designer (Graphic)
function getDesignerUserList() {
  const SHEET_ID = '1MKY52JEAtrCZJFnloyBN8cZMBXGOiDRF8rC-3ZFuC4U';
  const SHEET_NAME = 'DATA Password';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const designerUsers = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const username = row[1] ? String(row[1]).trim() : ''; // Col B
      const position = row[3] ? String(row[3]).trim().toLowerCase() : ''; // Col D

      // กรองเฉพาะตำแหน่งที่มีคำว่า "graphic" หรือ "กราฟิก"
      if (username && (position.includes('graphic') || position.includes('กราฟิก'))) {
        designerUsers.push(username);
      }
    }

    return designerUsers.sort();

  } catch (e) {
    Logger.log("Error getting designer users: " + e.message);
    return [];
  }
}

// ฟังก์ชันดึงรายชื่อลูกค้าจากชีต "รวมชื่อลูกค้า"
function getCustomerList() {
  const SHEET_ID = '1LLFrk7iebqHVutJKgIAhpAwBBwndRz_s6PaAuFRau0g'; // ID ที่คุณระบุ
  const SHEET_NAME = 'รวมชื่อลูกค้า';

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const customers = [];

    // เริ่มที่ i = 1 เพื่อข้าม Header (สมมติว่ารายชื่อลูกค้าอยู่คอลัมน์ B)
    for (let i = 1; i < data.length; i++) {
      const customerName = data[i][1] ? String(data[i][1]).trim() : '';
      if (customerName) {
        customers.push(customerName);
      }
    }
    return customers.sort(); // เรียงลำดับตามตัวอักษร
  } catch (e) {
    Logger.log("Error getting customer list: " + e.message);
    return [];
  }
}

// ==========================================
// ฟังก์ชันดึงรูปภาพเทมเพลตเป็น Base64
function getTemplateImageBase64() {
  const fileId = '1oV_F4XGdYq_t7CY6LE5jpN9HC_NETht8';
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);
    return 'data:' + blob.getContentType() + ';base64,' + base64;
  } catch (e) {
    Logger.log('Error getting template image: ' + e.message);
    return null;
  }
}
