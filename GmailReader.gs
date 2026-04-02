function getInvoiceEmails() {
  var label = GmailApp.getUserLabelByName(CONFIG.LABEL_INVOICE);
  if (!label) {
    Logger.log('label_not_found=' + CONFIG.LABEL_INVOICE);
    return [];
  }

  var maxThreads = parseInt(getScriptProperty_('MAX_THREADS', '5'), 10);
  var searchDays = parseInt(getScriptProperty_('SEARCH_DAYS', '7'), 10);
  var onlyLatest = String(getScriptProperty_('PROCESS_ONLY_LATEST_MESSAGE', 'true')).toLowerCase() === 'true';

  var query = [
    'label:' + CONFIG.LABEL_INVOICE,
    'newer_than:' + searchDays + 'd',
    '-label:' + CONFIG.LABEL_DONE,
    '-label:' + CONFIG.LABEL_MANUAL,
    '-label:' + CONFIG.LABEL_NO_PDF
  ].join(' ');

  var threads = GmailApp.search(query, 0, maxThreads);

  Logger.log('invoice_threads_found=' + threads.length);

  var entries = [];

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    if (!messages.length) return;

    if (onlyLatest) {
      entries.push({
        thread: thread,
        message: messages[messages.length - 1]
      });
      return;
    }

    messages.forEach(function(message) {
      entries.push({
        thread: thread,
        message: message
      });
    });
  });

  Logger.log('invoice_email_entries=' + entries.length);
  return entries;
}

function autoLabelInvoiceEmails() {
  var enabled = String(getScriptProperty_('AUTO_LABEL_ENABLED', 'true')).toLowerCase();
  if (enabled !== 'true') {
    Logger.log('auto_label_disabled');
    return;
  }

  var days = parseInt(getScriptProperty_('AUTO_LABEL_DAYS', '7'), 10);
  var maxThreads = parseInt(getScriptProperty_('AUTO_LABEL_MAX_THREADS', '10'), 10);

  var query = 'in:inbox newer_than:' + days + 'd has:attachment';
  var threads = GmailApp.search(query, 0, maxThreads);

  Logger.log('auto_label_threads_found=' + threads.length);

  var invoiceLabel = getOrCreateLabel_(CONFIG.LABEL_INVOICE);

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var matched = false;

    for (var i = 0; i < messages.length; i++) {
      if (isRealInvoiceEmail(messages[i])) {
        matched = true;
        break;
      }
    }

    if (matched) {
      thread.addLabel(invoiceLabel);
      Logger.log('auto_labeled_thread=' + thread.getId());
    }
  });
}

function isRealInvoiceEmail(message) {
  var subject = safeString_(message.getSubject()).toLowerCase();
  var from = safeString_(message.getFrom()).toLowerCase();
  var body = safeString_(message.getPlainBody()).toLowerCase();

  var text = [subject, from, body].join(' ');

  var blacklist = [
    'bhxh',
    'bhyt',
    'bhtn',
    'bảo hiểm',
    'bao hiem',
    'thông báo',
    'thong bao',
    'sao kê',
    'sao ke',
    'training',
    'otp',
    'mã xác thực',
    'ma xac thuc',
    'xác thực',
    'xac thuc',
    'khóa học',
    'khoa hoc',
    'đăng ký',
    'dang ky',
    'công văn',
    'cong van',
    'biên bản',
    'bien ban',
    'báo cáo',
    'bao cao'
  ];

  for (var i = 0; i < blacklist.length; i++) {
    if (text.indexOf(blacklist[i]) !== -1) {
      return false;
    }
  }

  var signals = [
    'hóa đơn',
    'hoa don',
    'invoice',
    'vat',
    'e-invoice',
    'einvoice',
    'mẫu số',
    'mau so',
    'ký hiệu',
    'ky hieu',
    'số hóa đơn',
    'so hoa don',
    'tra cứu hóa đơn',
    'tra cuu hoa don'
  ];

  var score = 0;
  for (var j = 0; j < signals.length; j++) {
    if (text.indexOf(signals[j]) !== -1) {
      score++;
    }
  }

  return score >= 2;
}

function getPdfAttachmentsFromMessage(message) {
  return getAttachmentsFromMessageByType_(message, function(contentType, name) {
    return contentType === 'application/pdf' || name.slice(-4) === '.pdf';
  });
}

function getXmlAttachmentsFromMessage(message) {
  return getAttachmentsFromMessageByType_(message, function(contentType, name) {
    return (
      contentType === 'text/xml' ||
      contentType === 'application/xml' ||
      name.slice(-4) === '.xml'
    );
  });
}

function getAttachmentsFromMessageByType_(message, matcher) {
  var attachments = message.getAttachments({ includeInlineImages: false });
  var filtered = [];

  attachments.forEach(function(att) {
    var contentType = safeString_(att.getContentType()).toLowerCase();
    var name = safeString_(att.getName()).toLowerCase();

    if (matcher(contentType, name)) {
      filtered.push(att);
    }
  });

  return filtered;
}

function getOrCreateLabel_(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

function ensureMainLabels_() {
  getOrCreateLabel_(CONFIG.LABEL_INVOICE);
  getOrCreateLabel_(CONFIG.LABEL_DONE);
  getOrCreateLabel_(CONFIG.LABEL_MANUAL);
  getOrCreateLabel_(CONFIG.LABEL_NO_PDF);
  getOrCreateLabel_(CONFIG.LABEL_ERROR);
}
