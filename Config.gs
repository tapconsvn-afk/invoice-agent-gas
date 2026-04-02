var CONFIG = {
  DRIVE_ROOT_FOLDER: 'Hoa_don_email_2026',
  SPREADSHEET_ID: '1uJgyyUUf2EiF4sFz4Zu6wuTFaeUl0ZABSN1TSEmMiIY',
  SHEET_DATA2: 'DATA2',
  SHEET_CHI_TIET_HOA_DON: 'CHI_TIET_HOA_DON',
  LABEL_INVOICE: 'Hoa_don_VAT',
  LABEL_DONE: 'HD_DA_XU_LY',
  LABEL_MANUAL: 'HD_CAN_TRA_TAY',
  LABEL_NO_PDF: 'HD_KHONG_TAI_DUOC_PDF',
  LABEL_ERROR: 'HD_LOI'
};

function getScriptProperty_(key, defaultValue) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  return value !== null ? value : defaultValue;
}

function setDefaultScriptProperties_() {
  var props = PropertiesService.getScriptProperties();

  var defaults = {
    AUTO_LABEL_ENABLED: 'true',
    AUTO_LABEL_DAYS: '7',
    AUTO_LABEL_MAX_THREADS: '10',
    MAX_THREADS: '5',
    SEARCH_DAYS: '7',
    OPENAI_ENABLED: 'true',
    OPENAI_MODEL: 'gpt-4o-mini',
    ENABLE_CHI_TIET: 'true',
    PROCESS_ONLY_LATEST_MESSAGE: 'true',
    MAX_RUNTIME_SECONDS: '240'
  };

  Object.keys(defaults).forEach(function(key) {
    if (props.getProperty(key) === null) {
      props.setProperty(key, defaults[key]);
    }
  });
}

function safeString_(value) {
  return value === undefined || value === null ? '' : String(value);
}

function normalizeKeyPart_(value) {
  return safeString_(value)
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '')
    .replace(/[|]/g, '');
}

function ensureDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number') {
    var d1 = new Date(value);
    if (!isNaN(d1.getTime())) return d1;
  }

  if (typeof value === 'string' && value) {
    var d2 = new Date(value);
    if (!isNaN(d2.getTime())) return d2;
  }

  return new Date();
}
