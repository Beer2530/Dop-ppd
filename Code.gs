/**
 * ============================================================================
 * ระบบหนังสือเวียนทราบ — กองการเตรียมพล กรมกำลังพลทหารบก (Code.gs)
 * Backend Google Apps Script (Google Sheets & Google Drive)
 * ============================================================================
 * วิธีติดตั้ง / อัปเดต:
 * 1. เปิด Google Spreadsheet ที่ต้องการเก็บข้อมูล
 * 2. ไปที่เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดนี้ทั้งหมดลงใน Code.gs
 * 4. กดบันทึก (รูปแผ่นดิสก์ หรือ Ctrl + S)
 * 5. กด การทำให้ใช้งานได้ (Deploy) > จัดการการทำให้ใช้งานได้ (Manage deployments)
 * 6. กด แก้ไข (ไอคอนรูปดินสอ)
 * 7. ที่ช่อง "เวอร์ชัน" (Version) ให้เลือก "เวอร์ชันใหม่" (New version) *** สำคัญมาก ***
 * 8. ที่ช่อง "ผู้มีสิทธิ์เข้าถึง" (Who has access) ให้เลือก "ทุกคน" (Anyone)
 * 9. กด ทำให้ใช้งานได้ (Deploy) แล้วคัดลอก Web App URL (ลงท้ายด้วย /exec)
 * ============================================================================
 */

var DOC_COLUMNS = [
  'id', 'title', 'description', 'fileName', 'driveFileId', 'mimeType',
  'fileSize', 'createdAt', 'deadline', 'status', 'closedAt', 'noReply',
  'replyCount', 'viewCount'
];

var REPLY_COLUMNS = [
  'id', 'docId', 'docTitle', 'name', 'choice', 'comment', 'createdAt'
];

var FOLDER_NAME = 'ระบบหนังสือเวียนทราบ_กพ.ทบ.';

// ----------------------------------------------------------------------------
// HTTP ENDPOINTS (GET & POST)
// ----------------------------------------------------------------------------

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';
  try {
    if (action === 'getState') {
      return jsonResponse_(getAppState());
    } else if (action === 'getPdf') {
      return jsonResponse_(getPdfBase64(params.fileId));
    }
    return jsonResponse_({
      success: true,
      message: 'Circular System API Active',
      timestamp: Date.now()
    });
  } catch (err) {
    return jsonResponse_({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // รอ Lock สูงสุด 30 วินาที เพื่อป้องกันการเขียนพร้อมกัน
    lock.waitLock(30000);
  } catch (lockErr) {
    return jsonResponse_({ success: false, error: 'ระบบไม่ว่าง กำลังมีผู้ใช้งานอื่นบันทึก กรุณาลองใหม่อีกครั้ง: ' + lockErr.toString() });
  }

  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var action = data.action || '';
    var result;

    if (action === 'getState') {
      result = getAppState();
    } else if (action === 'saveDoc') {
      result = saveDocument(data.doc);
    } else if (action === 'uploadChunk') {
      result = uploadFileChunk(data.docId, data.index, data.total, data.chunkBase64, data.fileName, data.mimeType);
    } else if (action === 'deleteDoc') {
      result = deleteDocument(data.docId);
    } else if (action === 'deleteAttachment') {
      result = deleteDocAttachmentOnly(data.docId);
    } else if (action === 'toggleStatus') {
      result = toggleDocStatus(data.docId, data.status);
    } else if (action === 'submitReply') {
      result = submitReply(data.reply);
    } else if (action === 'incrementView') {
      result = incrementDocView(data.docId);
    } else if (action === 'getPdf') {
      result = getPdfBase64(data.fileId);
    } else if (action === 'importDoc') {
      result = importDocument(data.doc);
    } else if (action === 'importReply') {
      result = importReply(data.reply);
    } else {
      result = { success: false, error: 'ไม่พบคำสั่งที่ระบุ: ' + action };
    }

    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ success: false, error: err.toString() });
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// DATABASE HELPER (SPREADSHEET SHEETS)
// ----------------------------------------------------------------------------

function getDb_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('ไม่พบ Active Spreadsheet กรุณาผูกสคริปต์นี้กับ Google Sheets');
  }

  var docSheet = ss.getSheetByName('Documents');
  if (!docSheet) {
    docSheet = ss.insertSheet('Documents');
    docSheet.appendRow(DOC_COLUMNS);
    docSheet.setFrozenRows(1);
    docSheet.getRange(1, 1, 1, DOC_COLUMNS.length).setFontWeight('bold').setBackground('#edf2e0');
  }

  var replySheet = ss.getSheetByName('Replies');
  if (!replySheet) {
    replySheet = ss.insertSheet('Replies');
    replySheet.appendRow(REPLY_COLUMNS);
    replySheet.setFrozenRows(1);
    replySheet.getRange(1, 1, 1, REPLY_COLUMNS.length).setFontWeight('bold').setBackground('#edf2e0');
  }

  return { ss: ss, docs: docSheet, replies: replySheet };
}

function getOrCreateFolder_() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
  }
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  return folder;
}

// ----------------------------------------------------------------------------
// CORE ACTIONS
// ----------------------------------------------------------------------------

/**
 * ดึงสถานะทั้งหมด (เอกสารเวียนทราบ และการตอบกลับ)
 */
function getAppState() {
  var db = getDb_();
  var docData = db.docs.getDataRange().getValues();
  var docs = [];
  if (docData.length > 1) {
    for (var i = 1; i < docData.length; i++) {
      var row = docData[i];
      if (!row[0]) continue;
      docs.push({
        id: String(row[0]),
        title: String(row[1] || ''),
        description: String(row[2] || ''),
        fileName: String(row[3] || ''),
        driveFileId: String(row[4] || ''),
        mimeType: String(row[5] || 'application/pdf'),
        fileSize: Number(row[6]) || 0,
        createdAt: Number(row[7]) || 0,
        deadline: Number(row[8]) || 0,
        status: String(row[9] || 'open'),
        closedAt: Number(row[10]) || 0,
        noReply: Boolean(row[11]),
        replyCount: Number(row[12]) || 0,
        viewCount: Number(row[13]) || 0
      });
    }
  }

  // เรียงลำดับตามวันที่สร้างล่าสุดอยู่บน
  docs.sort(function(a, b) {
    return (b.createdAt || 0) - (a.createdAt || 0);
  });

  var replyData = db.replies.getDataRange().getValues();
  var replies = [];
  if (replyData.length > 1) {
    for (var j = 1; j < replyData.length; j++) {
      var rRow = replyData[j];
      if (!rRow[0]) continue;
      replies.push({
        id: String(rRow[0]),
        docId: String(rRow[1] || ''),
        docTitle: String(rRow[2] || ''),
        name: String(rRow[3] || ''),
        choice: String(rRow[4] || ''),
        comment: String(rRow[5] || ''),
        createdAt: Number(rRow[6]) || 0
      });
    }
  }

  return { success: true, docs: docs, replies: replies };
}

/**
 * บันทึก หรือ แก้ไขข้อมูลหนังสือเวียนทราบ
 */
function saveDocument(doc) {
  if (!doc || !doc.id) {
    throw new Error('ข้อมูลหนังสือเวียนทราบไม่ถูกต้อง (Missing doc.id)');
  }
  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(doc.id)) {
      foundRow = i + 1; // 1-indexed
      break;
    }
  }

  var rowValues = [
    String(doc.id),
    String(doc.title || ''),
    String(doc.description || ''),
    String(doc.fileName || ''),
    String(doc.driveFileId || ''),
    String(doc.mimeType || 'application/pdf'),
    Number(doc.fileSize) || 0,
    Number(doc.createdAt) || Date.now(),
    Number(doc.deadline) || 0,
    String(doc.status || 'open'),
    Number(doc.closedAt) || 0,
    Boolean(doc.noReply),
    Number(doc.replyCount) || 0,
    Number(doc.viewCount) || 0
  ];

  if (foundRow > 0) {
    // ถ้ารายการเดิมมี driveFileId อยู่แล้ว และข้อมูลที่ส่งมาไม่ได้สั่งลบ (fileName ยังมี) ให้เก็บ driveFileId เดิมไว้
    var oldDriveFileId = String(data[foundRow - 1][4] || '');
    if (oldDriveFileId && !rowValues[4] && rowValues[3]) {
      rowValues[4] = oldDriveFileId;
    }
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { success: true, docId: doc.id };
}

/**
 * อัปโหลดไฟล์แบบ Chunked (ทีละชิ้น) และรวมไฟล์บันทึกเข้า Google Drive อย่างถาวร
 */
function uploadFileChunk(docId, index, total, chunkBase64, fileName, mimeType) {
  if (!docId) throw new Error('Missing docId');
  var cache = CacheService.getScriptCache();
  var cacheKey = 'chk_' + docId + '_' + index;

  // บันทึกชิ้นส่วนลง CacheService (อยู่ได้ 6 ชั่วโมง)
  cache.put(cacheKey, chunkBase64, 21600);

  // ถ้ายังไม่ใช่ชิ้นสุดท้าย ให้ตอบรับชิ้นส่วน
  if (Number(index) < Number(total) - 1) {
    return { success: true, completed: false, index: index, total: total };
  }

  // ชิ้นสุดท้าย: รวบรวมชิ้นส่วนทั้งหมดเข้าด้วยกัน
  var fullBase64 = '';
  for (var i = 0; i < total; i++) {
    var k = 'chk_' + docId + '_' + i;
    var piece = (i === index) ? chunkBase64 : cache.get(k);
    if (!piece) {
      throw new Error('ชิ้นส่วนไฟล์ที่ ' + (i + 1) + ' ขาดหาย กรุณาอัปโหลดไฟล์ใหม่อีกครั้ง');
    }
    fullBase64 += piece;
  }

  var bytes = Utilities.base64Decode(fullBase64);
  var blob = Utilities.newBlob(bytes, mimeType || 'application/pdf', fileName || ('doc_' + docId + '.pdf'));

  // บันทึกไฟล์ลง Google Drive
  var folder = getOrCreateFolder_();
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  var fileId = file.getId();
  var fileSize = file.getSize();

  // ล้างชิ้นส่วนในแคช
  for (var c = 0; c < total; c++) {
    cache.remove('chk_' + docId + '_' + c);
  }

  // อัปเดตข้อมูลแถวใน Google Sheets ให้มี driveFileId, fileName, fileSize ทันที
  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(docId)) {
      sheet.getRange(r + 1, 4).setValue(fileName); // fileName (col 4)
      sheet.getRange(r + 1, 5).setValue(fileId);   // driveFileId (col 5)
      sheet.getRange(r + 1, 6).setValue(mimeType || 'application/pdf'); // mimeType (col 6)
      sheet.getRange(r + 1, 7).setValue(fileSize); // fileSize (col 7)
      break;
    }
  }

  return {
    success: true,
    completed: true,
    driveFileId: fileId,
    fileName: fileName,
    fileSize: fileSize
  };
}

/**
 * ดึงข้อมูลไฟล์ PDF จาก Google Drive ส่งกลับเป็น Base64
 */
function getPdfBase64(fileId) {
  if (!fileId) throw new Error('ไม่ระบุ File ID');
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var bytes = blob.getBytes();
  var b64 = Utilities.base64Encode(bytes);
  return {
    success: true,
    base64: b64,
    fileName: file.getName(),
    fileSize: file.getSize(),
    mimeType: blob.getContentType() || 'application/pdf'
  };
}

/**
 * ลบเฉพาะไฟล์แนบออกจาก Google Drive และล้างค่าในชีต
 */
function deleteDocAttachmentOnly(docId) {
  if (!docId) throw new Error('Missing docId');
  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(docId)) {
      var driveFileId = String(data[i][4] || '');
      if (driveFileId) {
        try {
          DriveApp.getFileById(driveFileId).setTrashed(true);
        } catch (e) {}
      }
      sheet.getRange(i + 1, 4).setValue(''); // fileName
      sheet.getRange(i + 1, 5).setValue(''); // driveFileId
      sheet.getRange(i + 1, 7).setValue(0);  // fileSize
      break;
    }
  }
  return { success: true };
}

/**
 * ลบหนังสือเวียนทราบ พร้อมลบไฟล์ใน Drive และคำตอบทั้งหมดที่เกี่ยวข้อง
 */
function deleteDocument(docId) {
  if (!docId) throw new Error('Missing docId');
  var db = getDb_();

  // 1. ลบจาก Documents
  var docSheet = db.docs;
  var docData = docSheet.getDataRange().getValues();
  for (var i = 1; i < docData.length; i++) {
    if (String(docData[i][0]) === String(docId)) {
      var driveFileId = String(docData[i][4] || '');
      if (driveFileId) {
        try {
          DriveApp.getFileById(driveFileId).setTrashed(true);
        } catch (e) {}
      }
      docSheet.deleteRow(i + 1);
      break;
    }
  }

  // 2. ลบคำตอบจาก Replies
  var replySheet = db.replies;
  var replyData = replySheet.getDataRange().getValues();
  for (var j = replyData.length - 1; j >= 1; j--) {
    if (String(replyData[j][1]) === String(docId)) {
      replySheet.deleteRow(j + 1);
    }
  }

  return { success: true };
}

/**
 * สลับสถานะเปิด/ปิดรับคำตอบ
 */
function toggleDocStatus(docId, newStatus) {
  if (!docId) throw new Error('Missing docId');
  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(docId)) {
      sheet.getRange(i + 1, 10).setValue(newStatus); // status (col 10)
      if (newStatus === 'closed') {
        sheet.getRange(i + 1, 11).setValue(Date.now()); // closedAt (col 11)
      }
      break;
    }
  }
  return { success: true };
}

/**
 * ส่งการตอบกลับ
 */
function submitReply(reply) {
  if (!reply || !reply.docId || !reply.name) {
    throw new Error('ข้อมูลการตอบกลับไม่ครบถ้วน');
  }
  var db = getDb_();
  var replyId = reply.id || ('rep-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6));

  db.replies.appendRow([
    replyId,
    String(reply.docId),
    String(reply.docTitle || ''),
    String(reply.name || ''),
    String(reply.choice || ''),
    String(reply.comment || ''),
    Number(reply.createdAt) || Date.now()
  ]);

  // เพิ่ม replyCount ใน Documents
  var docSheet = db.docs;
  var docData = docSheet.getDataRange().getValues();
  for (var i = 1; i < docData.length; i++) {
    if (String(docData[i][0]) === String(reply.docId)) {
      var curr = Number(docData[i][12]) || 0;
      docSheet.getRange(i + 1, 13).setValue(curr + 1);
      break;
    }
  }

  return { success: true, replyId: replyId };
}

/**
 * เพิ่มจำนวนครั้งที่เปิดดู
 */
function incrementDocView(docId) {
  if (!docId) return { success: false };
  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(docId)) {
      var curr = Number(data[i][13]) || 0;
      sheet.getRange(i + 1, 14).setValue(curr + 1);
      break;
    }
  }
  return { success: true };
}

/**
 * นำเข้า Document (Batch import)
 */
function importDocument(doc) {
  return saveDocument(doc);
}

/**
 * นำเข้า Reply (Batch import)
 */
function importReply(reply) {
  if (!reply || !reply.id) return { success: false };
  var db = getDb_();
  var sheet = db.replies;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(reply.id)) {
      return { success: true, existing: true };
    }
  }
  sheet.appendRow([
    String(reply.id),
    String(reply.docId || ''),
    String(reply.docTitle || ''),
    String(reply.name || ''),
    String(reply.choice || ''),
    String(reply.comment || ''),
    Number(reply.createdAt) || Date.now()
  ]);
  return { success: true };
}
