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
 * 9. กด ทำให้ใช้งานได้ (Deploy)
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

var TARGET_FOLDER_ID = '1G6MroXy3X_nf3Fly3VvkjP0FDnbYMh4g';
var FOLDER_NAME = 'ระบบหนังสือเวียนทราบ_กพ.ทบ.';
var TMP_FOLDER_NAME = '_upload_tmp';

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
    } else if (action === 'listDriveFiles') {
      return jsonResponse_(listDriveFiles());
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
    } else if (action === 'listDriveFiles') {
      result = listDriveFiles();
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
    } else {
      result = { success: false, error: 'ไม่พบคำสั่งที่ระบุ: ' + action };
    }

    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ success: false, error: err.toString() });
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {}
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// DATABASE & DRIVE HELPERS
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
  var folder = null;
  if (TARGET_FOLDER_ID) {
    try {
      folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    } catch (err) {
      folder = null;
    }
  }
  if (!folder) {
    var folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(FOLDER_NAME);
    }
  }
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  return folder;
}

function getOrCreateTempFolder_() {
  var rootFolder = getOrCreateFolder_();
  var subs = rootFolder.getFoldersByName(TMP_FOLDER_NAME);
  if (subs.hasNext()) {
    return subs.next();
  }
  return rootFolder.createFolder(TMP_FOLDER_NAME);
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

  docs.sort(function(a, b) {
    return (b.createdAt || 0) - (a.createdAt || 0);
  });

  var replyData = db.replies.getDataRange().getValues();
  var replies = [];
  if (replyData.length > 1) {
    for (var j = 1; j < replyData.length; j++) {
      var rRow = replyData[j];
      if (!rRow[0]) continue;
      var choiceVal = String(rRow[4] || '');
      var tsVal = Number(rRow[6]) || 0;
      replies.push({
        id: String(rRow[0]),
        docId: String(rRow[1] || ''),
        docTitle: String(rRow[2] || ''),
        name: String(rRow[3] || ''),
        choice: choiceVal,
        response: choiceVal,
        comment: String(rRow[5] || ''),
        createdAt: tsVal,
        ts: tsVal
      });
    }
  }

  return { success: true, docs: docs, replies: replies };
}

/**
 * ดึงรายการไฟล์ PDF ปัจจุบันจากโฟลเดอร์ Google Drive หลังบ้าน และ PDF ล่าสุดใน Drive
 */
function listDriveFiles() {
  var folder = getOrCreateFolder_();
  var files = [];
  var seenIds = {};

  var folderIter = folder.getFiles();
  while (folderIter.hasNext()) {
    var f = folderIter.next();
    if (f.isTrashed()) continue;
    var name = f.getName();
    var mime = f.getMimeType() || '';
    if (mime !== 'application/pdf' && name.toLowerCase().slice(-4) !== '.pdf') continue;
    var id = f.getId();
    seenIds[id] = true;
    files.push({
      id: id,
      name: name,
      size: f.getSize(),
      mimeType: mime || 'application/pdf',
      updatedAt: f.getLastUpdated() ? f.getLastUpdated().getTime() : 0,
      inSystemFolder: true
    });
  }

  try {
    var allPdfIter = DriveApp.getFilesByType(MimeType.PDF);
    var extraCount = 0;
    while (allPdfIter.hasNext() && extraCount < 30) {
      var pf = allPdfIter.next();
      if (pf.isTrashed()) continue;
      var pid = pf.getId();
      if (seenIds[pid]) continue;
      var pname = pf.getName() || '';
      if (pname.slice(-5) === '.part') continue;
      seenIds[pid] = true;
      extraCount++;
      files.push({
        id: pid,
        name: pname,
        size: pf.getSize(),
        mimeType: pf.getMimeType() || 'application/pdf',
        updatedAt: pf.getLastUpdated() ? pf.getLastUpdated().getTime() : 0,
        inSystemFolder: false
      });
    }
  } catch (e) {}

  files.sort(function(a, b) {
    if (a.inSystemFolder !== b.inSystemFolder) return a.inSystemFolder ? -1 : 1;
    return (b.updatedAt || 0) - (a.updatedAt || 0);
  });

  return {
    success: true,
    folderId: folder.getId(),
    folderName: folder.getName(),
    folderUrl: 'https://drive.google.com/drive/folders/' + folder.getId(),
    files: files
  };
}

/**
 * บันทึก หรือ แก้ไขข้อมูลหนังสือเวียนทราบ (และนำไฟล์เข้าโฟลเดอร์หลังบ้านให้อัตโนมัติ)
 */
function saveDocument(doc) {
  if (!doc || !doc.id) {
    throw new Error('ข้อมูลหนังสือเวียนทราบไม่ถูกต้อง (Missing doc.id)');
  }

  var driveFileId = String(doc.driveFileId || '');
  var fileName = String(doc.fileName || '');
  var fileSize = Number(doc.fileSize) || 0;
  var mimeType = String(doc.mimeType || 'application/pdf');

  if (driveFileId) {
    try {
      var targetFolder = getOrCreateFolder_();
      var driveFile = DriveApp.getFileById(driveFileId);
      try {
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {}
      try {
        var inTarget = false;
        var parents = driveFile.getParents();
        while (parents.hasNext()) {
          if (parents.next().getId() === targetFolder.getId()) {
            inTarget = true;
            break;
          }
        }
        if (!inTarget) {
          driveFile.moveTo(targetFolder);
        }
      } catch (moveErr) {}
      if (!fileName) fileName = driveFile.getName();
      if (!fileSize) fileSize = driveFile.getSize();
      mimeType = driveFile.getMimeType() || 'application/pdf';
    } catch (fileErr) {
      throw new Error('ไม่พบไฟล์ใน Google Drive (ID: ' + driveFileId + '): ' + fileErr.toString());
    }
  }

  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(doc.id)) {
      foundRow = i + 1;
      break;
    }
  }

  var rowValues = [
    String(doc.id),
    String(doc.title || ''),
    String(doc.description || ''),
    fileName,
    driveFileId,
    mimeType,
    fileSize,
    Number(doc.createdAt) || Date.now(),
    Number(doc.deadline) || 0,
    String(doc.status || 'open'),
    Number(doc.closedAt) || 0,
    Boolean(doc.noReply),
    Number(doc.replyCount) || 0,
    Number(doc.viewCount) || 0
  ];

  if (foundRow > 0) {
    var oldDriveFileId = String(data[foundRow - 1][4] || '');
    if (oldDriveFileId && !rowValues[4] && rowValues[3]) {
      rowValues[4] = oldDriveFileId;
    }
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { success: true, docId: doc.id, driveFileId: rowValues[4], fileName: rowValues[3], fileSize: rowValues[6] };
}

/**
 * อัปโหลดไฟล์เข้า Google Drive ถาวร (ไม่มีขีดจำกัด CacheService 100KB และไม่มีปัญหา Base64 padding)
 */
function uploadFileChunk(docId, index, total, chunkBase64, fileName, mimeType) {
  if (!docId) throw new Error('Missing docId');
  var idx = Number(index) || 0;
  var tot = Math.max(1, Number(total) || 1);
  var safeName = fileName || ('doc_' + docId + '.pdf');
  var safeMime = mimeType || 'application/pdf';
  var folder = getOrCreateFolder_();
  var file;

  if (tot === 1) {
    // อัปโหลดรวดเดียวเสร็จทันที (ไม่ผ่าน CacheService)
    var singleBytes = Utilities.base64Decode(chunkBase64);
    var singleBlob = Utilities.newBlob(singleBytes, safeMime, safeName);
    file = folder.createFile(singleBlob);
  } else {
    // กรณีไฟล์ใหญ่มาก แบ่งเก็บชิ้นส่วนไบนารีในโฟลเดอร์ชั่วคราว _upload_tmp
    var tmpFolder = getOrCreateTempFolder_();
    var partName = docId + '_' + idx + '.part';
    var existingParts = tmpFolder.getFilesByName(partName);
    while (existingParts.hasNext()) {
      existingParts.next().setTrashed(true);
    }

    var partBytes = Utilities.base64Decode(chunkBase64);
    tmpFolder.createFile(Utilities.newBlob(partBytes, 'application/octet-stream', partName));

    if (idx < tot - 1) {
      return { success: true, completed: false, index: idx, total: tot };
    }

    // ชิ้นสุดท้าย: รวมไบต์จากทุกชิ้นส่วนตามลำดับ
    var allBytes = [];
    var trashList = [];
    for (var i = 0; i < tot; i++) {
      var pName = docId + '_' + i + '.part';
      var pIter = tmpFolder.getFilesByName(pName);
      if (!pIter.hasNext()) {
        throw new Error('ชิ้นส่วนไฟล์ที่ ' + (i + 1) + ' ขาดหาย กรุณาอัปโหลดไฟล์ใหม่อีกครั้ง');
      }
      var pFile = pIter.next();
      var pBytes = pFile.getBlob().getBytes();
      for (var b = 0; b < pBytes.length; b++) {
        allBytes.push(pBytes[b]);
      }
      trashList.push(pFile);
    }

    var finalBlob = Utilities.newBlob(allBytes, safeMime, safeName);
    file = folder.createFile(finalBlob);

    for (var t = 0; t < trashList.length; t++) {
      try {
        trashList[t].setTrashed(true);
      } catch (e) {}
    }
  }

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  var fileId = file.getId();
  var fileSize = file.getSize();

  var db = getDb_();
  var sheet = db.docs;
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(docId)) {
      sheet.getRange(r + 1, 4).setValue(safeName);
      sheet.getRange(r + 1, 5).setValue(fileId);
      sheet.getRange(r + 1, 6).setValue(safeMime);
      sheet.getRange(r + 1, 7).setValue(fileSize);
      break;
    }
  }

  return {
    success: true,
    completed: true,
    driveFileId: fileId,
    fileName: safeName,
    fileSize: fileSize
  };
}

/**
 * ดึงข้อมูลไฟล์ PDF จาก Google Drive ส่งกลับเป็น Base64 (รองรับ fallback กรณีไฟล์ใหญ่เกินขีดจำกัด GAS response)
 */
function getPdfBase64(fileId) {
  if (!fileId) throw new Error('ไม่ระบุ File ID');
  var file = DriveApp.getFileById(fileId);
  var fileSize = file.getSize();
  var fileName = file.getName();

  // หากไฟล์เกิน 26 MB ส่งสัญญาณให้หน้าเว็บใช้ Google Drive Preview + Download โดยตรง
  if (fileSize > 26 * 1024 * 1024) {
    return {
      success: true,
      tooLargeForBase64: true,
      fileName: fileName,
      fileSize: fileSize,
      mimeType: 'application/pdf'
    };
  }

  var blob = file.getBlob();
  var bytes = blob.getBytes();
  var b64 = Utilities.base64Encode(bytes);
  return {
    success: true,
    base64: b64,
    fileName: fileName,
    fileSize: fileSize,
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
      sheet.getRange(i + 1, 4).setValue('');
      sheet.getRange(i + 1, 5).setValue('');
      sheet.getRange(i + 1, 7).setValue(0);
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
      sheet.getRange(i + 1, 10).setValue(newStatus);
      if (newStatus === 'closed') {
        sheet.getRange(i + 1, 11).setValue(Date.now());
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
  var choiceText = String(reply.choice || reply.response || '');
  var createdTs = Number(reply.createdAt || reply.ts) || Date.now();

  db.replies.appendRow([
    replyId,
    String(reply.docId),
    String(reply.docTitle || ''),
    String(reply.name || ''),
    choiceText,
    String(reply.comment || ''),
    createdTs
  ]);

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
