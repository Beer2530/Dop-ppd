function importFromOldApp() {
  const EXPORT_URL = 'https://ais-dev-qbit3afohggyyrkyzrjzv4-343387795750.asia-east1.run.app/api/export';
  const res = UrlFetchApp.fetch(EXPORT_URL, { muteHttpExceptions: true });
  const payload = JSON.parse(res.getContentText());

  const { docSheet, replySheet, folder } = getDb_();
  const existingDocIds = new Set(docSheet.getDataRange().getValues().slice(1).map(r => String(r[0])));
  const existingReplyIds = new Set(replySheet.getDataRange().getValues().slice(1).map(r => String(r[0])));

  // 1. นำเข้าหนังสือเวียนทราบ + ดึงไฟล์ PDF เข้า Drive ใหม่
  (payload.docs || []).forEach((d) => {
    if (existingDocIds.has(String(d.id))) return;

    let newDriveId = '';
    if (d.pdfDownloadUrl && d.hasFileOnServer) {
      try {
        const pdfRes = UrlFetchApp.fetch(d.pdfDownloadUrl, { muteHttpExceptions: true });
        if (pdfRes.getResponseCode() === 200) {
          const blob = pdfRes.getBlob().setName(d.fileName || (d.title + '.pdf'));
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          newDriveId = file.getId();
        }
      } catch (e) {}
    }

    docSheet.appendRow([
      d.id, d.title || '', d.description || '',
      newDriveId ? (d.fileName || '') : '',
      d.mimeType || 'application/pdf',
      newDriveId ? (Number(d.fileSize) || 0) : 0,
      newDriveId,
      Number(d.createdAt) || Date.now(),
      Number(d.deadline) || 0,
      d.status || 'open',
      Number(d.closedAt) || 0,
      Boolean(d.noReply),
      Number(d.replyCount) || 0,
      Number(d.viewCount) || 0
    ]);
  });

  // 2. นำเข้าคำตอบกลับทั้งหมด
  (payload.replies || []).forEach((r) => {
    if (existingReplyIds.has(String(r.id))) return;
    replySheet.appendRow([
      r.id, r.docId || '', r.docTitle || '', r.name || '',
      r.response || '', r.comment || '',
      Number(r.ts) || Date.now(), r.timeFormatted || ''
    ]);
  });
}