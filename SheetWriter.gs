function ensureMainSheets_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var data2 = ss.getSheetByName(CONFIG.SHEET_DATA2);
  if (!data2) {
    data2 = ss.insertSheet(CONFIG.SHEET_DATA2);
  }

  var detail = ss.getSheetByName(CONFIG.SHEET_CHI_TIET_HOA_DON);
  if (!detail) {
    detail = ss.insertSheet(CONFIG.SHEET_CHI_TIET_HOA_DON);
  }

  var data2Headers = [[
    'Thoi_gian_xu_ly',
    'Ngay_email',
    'Nguoi_gui',
    'Tieu_de',
    'Ten_file_pdf',
    'Link_pdf',
    'Ten_file_xml',
    'Link_xml',
    'So_hoa_don',
    'Ky_hieu',
    'Ngay_hoa_don',
    'MST_ben_ban',
    'Nguoi_ban',
    'Tien_truoc_thue',
    'Tien_thue_GTGT',
    'Tong_tien',
    'Trang_thai',
    'Ghi_chu',
    'Khoa_chong_trung'
  ]];

  var detailHeaders = [[
    'Thoi_gian_xu_ly',
    'Khoa_chong_trung',
    'So_hoa_don',
    'Ky_hieu',
    'Ngay_hoa_don',
    'MST_ben_ban',
    'Nguoi_ban',
    'STT_dong',
    'Ten_hang_hoa',
    'Don_vi_tinh',
    'So_luong',
    'Don_gia',
    'Thanh_tien',
    'Thue_suat',
    'Tien_thue',
    'Ghi_chu'
  ]];

  if (data2.getLastRow() === 0) {
    data2.getRange(1, 1, 1, data2Headers[0].length).setValues(data2Headers);
  }

  if (detail.getLastRow() === 0) {
    detail.getRange(1, 1, 1, detailHeaders[0].length).setValues(detailHeaders);
  }
}

function appendInvoiceRow(data) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_DATA2);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + CONFIG.SHEET_DATA2);
  }

  sheet.appendRow([
    data.thoiGianXuLy,
    data.ngayEmail,
    data.nguoiGui,
    data.tieuDe,
    data.tenFilePdf,
    data.linkPdf,
    data.tenFileXml,
    data.linkXml,
    data.soHoaDon,
    data.kyHieu,
    data.ngayHoaDon,
    data.mstBenBan,
    data.nguoiBan,
    data.tienTruocThue,
    data.tienThueGtgt,
    data.tongTien,
    data.trangThai,
    data.ghiChu,
    data.khoaChongTrung
  ]);
}

function appendInvoiceItems(duplicateKey, items, meta) {
  var enabled = String(getScriptProperty_('ENABLE_CHI_TIET', 'true')).toLowerCase();
  if (enabled !== 'true') return;

  if (!items || !items.length) return;

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_CHI_TIET_HOA_DON);

  if (!sheet) {
    Logger.log('sheet_not_found=' + CONFIG.SHEET_CHI_TIET_HOA_DON);
    return;
  }

  var now = new Date();
  var rows = items.map(function(item, index) {
    return [
      now,
      duplicateKey,
      safeString_(meta.soHoaDon),
      safeString_(meta.kyHieu),
      safeString_(meta.ngayHoaDon),
      safeString_(meta.mstBenBan),
      safeString_(meta.nguoiBan),
      index + 1,
      safeString_(item.ten_hang_hoa),
      safeString_(item.don_vi_tinh),
      safeString_(item.so_luong),
      safeString_(item.don_gia),
      safeString_(item.thanh_tien),
      safeString_(item.thue_suat),
      safeString_(item.tien_thue),
      safeString_(item.ghi_chu)
    ];
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function loadDuplicateKeySet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_DATA2);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + CONFIG.SHEET_DATA2);
  }

  var set = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return set;

  var values = sheet.getRange(2, 19, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = safeString_(values[i][0]).trim();
    if (key) set[key] = true;
  }

  return set;
}

function isDuplicateKeyExists_(duplicateKey) {
  if (!duplicateKey) return false;

  var existing = loadDuplicateKeySet_();
  return !!existing[safeString_(duplicateKey).trim()];
}
