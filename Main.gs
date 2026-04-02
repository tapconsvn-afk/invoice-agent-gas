function processInvoiceEmails() {
  Logger.log('process_started');

  setDefaultScriptProperties_();
  autoLabelInvoiceEmails();

  var emailEntries = getInvoiceEmails();
  var duplicateSet = loadDuplicateKeySet_();
  var startedAt = new Date().getTime();
  var maxRuntimeMs = parseInt(getScriptProperty_('MAX_RUNTIME_SECONDS', '240'), 10) * 1000;

  Logger.log('invoice_email_entries=' + emailEntries.length);

  var processed = 0;
  var skipped = 0;
  var errors = 0;

  for (var i = 0; i < emailEntries.length; i++) {
    if (new Date().getTime() - startedAt > maxRuntimeMs) {
      Logger.log('runtime_limit_reached index=' + i);
      break;
    }

    var entry = emailEntries[i];

    try {
      var result = processOneInvoiceEmail_(entry, duplicateSet);

      if (result === 'processed') {
        processed++;
      } else if (result === 'skipped') {
        skipped++;
      } else {
        skipped++;
      }
    } catch (err) {
      errors++;
      Logger.log('process_error=' + err.message);

      try {
        markEmailError_(entry.thread, entry.message, err);
      } catch (labelErr) {
        Logger.log('label_error=' + labelErr.message);
      }
    }
  }

  Logger.log('process_done processed=' + processed + ', skipped=' + skipped + ', errors=' + errors);
}

function processOneInvoiceEmail_(entry, duplicateSet) {
  var message = entry.message;
  var thread = entry.thread;

  var messageId = safeString_(message.getId());
  var subject = safeString_(message.getSubject());
  var from = safeString_(message.getFrom());
  var emailDate = ensureDate_(message.getDate());

  Logger.log('processing_message=' + messageId);

  var precheckKey = buildDuplicateKey_(messageId, subject, from, '', '', '');
  if (isDuplicateKeyInSet_(precheckKey, duplicateSet)) {
    Logger.log('duplicate_message_skip=' + precheckKey);
    return 'skipped';
  }

  var pdfs = getPdfAttachmentsFromMessage(message);
  var xmls = getXmlAttachmentsFromMessage(message);

  var pdfFile = null;
  var pdfUrl = '';
  var pdfName = '';

  if (pdfs.length > 0) {
    pdfFile = saveAttachmentToMonthlyFolder_(pdfs[0], emailDate);
    pdfUrl = pdfFile ? pdfFile.getUrl() : '';
    pdfName = pdfFile ? pdfFile.getName() : '';
    Logger.log('saved_file=' + pdfName);
  }

  var extracted = buildEmptyInvoiceResult_();
  if (pdfFile) {
    extracted = extractInvoiceDataFromPdf(pdfFile);
  }

  var invoiceNumber = getField_(extracted, 'so_hoa_don');
  var invoiceSymbol = getField_(extracted, 'ky_hieu');
  var invoiceDate = getField_(extracted, 'ngay_hoa_don');
  var sellerTaxCode = getField_(extracted, 'mst_ben_ban');
  var sellerName = getField_(extracted, 'nguoi_ban');
  var beforeTax = getField_(extracted, 'tien_truoc_thue');
  var vatAmount = getField_(extracted, 'tien_thue_gtgt');
  var totalAmount = getField_(extracted, 'tong_tien');
  var items = getItems_(extracted);

  var duplicateKey = buildDuplicateKey_(messageId, subject, from, invoiceNumber, invoiceSymbol, sellerTaxCode);
  Logger.log('duplicate_key=' + duplicateKey);

  if (isDuplicateKeyInSet_(duplicateKey, duplicateSet)) {
    Logger.log('duplicate_invoice_skip=' + duplicateKey);
    return 'skipped';
  }

  var status = '';
  var note = '';

  if (pdfFile) {
    if (hasEnoughInvoiceData_(invoiceNumber, sellerName, totalAmount, beforeTax, vatAmount)) {
      status = 'DA_XU_LY_OK';
      note = 'Đã lưu PDF và trích xuất được dữ liệu chính';
      markEmailDone_(thread);
    } else {
      status = 'CAN_TRA_TAY';
      note = 'Có PDF nhưng dữ liệu trích xuất chưa đủ, cần kiểm tra tay';
      markEmailManual_(thread);
    }
  } else {
    status = 'KHONG_TAI_DUOC_PDF';
    note = buildNoPdfNote_(xmls, subject);
    markEmailNoPdf_(thread);
  }

  appendInvoiceRow({
    thoiGianXuLy: new Date(),
    ngayEmail: emailDate,
    nguoiGui: from,
    tieuDe: subject,
    tenFilePdf: pdfName,
    linkPdf: pdfUrl,
    tenFileXml: '',
    linkXml: '',
    soHoaDon: invoiceNumber,
    kyHieu: invoiceSymbol,
    ngayHoaDon: invoiceDate,
    mstBenBan: sellerTaxCode,
    nguoiBan: sellerName,
    tienTruocThue: beforeTax,
    tienThueGtgt: vatAmount,
    tongTien: totalAmount,
    trangThai: status,
    ghiChu: note,
    khoaChongTrung: duplicateKey
  });

  duplicateSet[duplicateKey] = true;
  duplicateSet[precheckKey] = true;

  appendInvoiceItems(duplicateKey, items, {
    soHoaDon: invoiceNumber,
    kyHieu: invoiceSymbol,
    ngayHoaDon: invoiceDate,
    mstBenBan: sellerTaxCode,
    nguoiBan: sellerName
  });

  return 'processed';
}

function isDuplicateKeyInSet_(duplicateKey, duplicateSet) {
  var key = safeString_(duplicateKey).trim();
  if (!key) return false;

  return !!(duplicateSet && duplicateSet[key]);
}

function buildDuplicateKey_(messageId, subject, from, invoiceNumber, invoiceSymbol, sellerTaxCode) {
  var so = normalizeKeyPart_(invoiceNumber);
  var ky = normalizeKeyPart_(invoiceSymbol);
  var mst = normalizeKeyPart_(sellerTaxCode);

  if (so) {
    return [so, ky, mst].join('|');
  }

  return [
    normalizeKeyPart_(messageId),
    normalizeKeyPart_(subject),
    normalizeKeyPart_(from)
  ].join('|');
}

function buildNoPdfNote_(xmls, subject) {
  if (xmls && xmls.length > 0) {
    return 'Email không có PDF, chỉ có XML hoặc file khác; cần tra tay';
  }

  if (safeString_(subject)) {
    return 'Không có PDF đính kèm; cần tra tay từ nội dung email';
  }

  return 'Không có PDF; cần tra tay';
}

function hasEnoughInvoiceData_(invoiceNumber, sellerName, totalAmount, beforeTax, vatAmount) {
  var count = 0;

  if (invoiceNumber) count++;
  if (sellerName) count++;
  if (totalAmount) count++;
  if (beforeTax) count++;
  if (vatAmount) count++;

  return count >= 3;
}

function markEmailDone_(thread) {
  markEmailStatus_(thread, CONFIG.LABEL_DONE);
}

function markEmailManual_(thread) {
  markEmailStatus_(thread, CONFIG.LABEL_MANUAL);
}

function markEmailNoPdf_(thread) {
  markEmailStatus_(thread, CONFIG.LABEL_NO_PDF);
}

function markEmailStatus_(thread, labelName) {
  removeProcessingLabels_(thread);
  thread.addLabel(getOrCreateLabel_(labelName));
}

function markEmailError_(thread, message, err) {
  markEmailStatus_(thread, CONFIG.LABEL_ERROR);
  Logger.log('mark_error_message=' + safeString_(message.getId()) + ' | ' + safeString_(err.message));
}

function removeProcessingLabels_(thread) {
  var names = [
    CONFIG.LABEL_DONE,
    CONFIG.LABEL_MANUAL,
    CONFIG.LABEL_NO_PDF,
    CONFIG.LABEL_ERROR
  ];

  for (var i = 0; i < names.length; i++) {
    var label = GmailApp.getUserLabelByName(names[i]);
    if (label) {
      thread.removeLabel(label);
    }
  }
}

function getField_(obj, key) {
  if (!obj || typeof obj !== 'object') return '';
  return obj[key] === undefined || obj[key] === null ? '' : obj[key];
}

function getItems_(obj) {
  if (!obj || typeof obj !== 'object') return [];
  if (!obj.items || !Array.isArray(obj.items)) return [];
  return obj.items;
}
