function extractInvoiceDataFromPdf(pdfFile) {
  return extractInvoiceDataWithOpenAI(pdfFile);
}

function analyzeInvoicePdfWithOpenAI(pdfFile) {
  return extractInvoiceDataWithOpenAI(pdfFile);
}

function extractInvoiceDataWithOpenAI(pdfFile) {
  var enabled = String(getScriptProperty_('OPENAI_ENABLED', 'true')).toLowerCase();
  if (enabled !== 'true') {
    Logger.log('openai_disabled');
    return buildEmptyInvoiceResult_();
  }

  var apiKey = getScriptProperty_('OPENAI_API_KEY', '');
  if (!apiKey) {
    Logger.log('openai_api_key_missing');
    return buildEmptyInvoiceResult_();
  }

  var model = getScriptProperty_('OPENAI_MODEL', 'gpt-4o-mini');
  var pdfText = extractTextFromPdfWithOcr_(pdfFile);

  if (!pdfText) {
    Logger.log('pdf_text_empty');
    return buildEmptyInvoiceResult_();
  }

  var prompt = buildInvoiceExtractionPrompt_(pdfText);
  var rawText = callOpenAIForInvoiceJson_(apiKey, model, prompt);
  var parsed = parseInvoiceJsonSafe_(rawText);

  if (!parsed) {
    Logger.log('openai_parse_failed');
    return buildEmptyInvoiceResult_();
  }

  return normalizeInvoiceResult_(parsed);
}

function extractTextFromPdfWithOcr_(pdfFile) {
  if (!pdfFile) return '';

  var tempDocId = '';

  try {
    var blob = pdfFile.getBlob();

    var resource = {
      title: 'TMP_OCR_' + new Date().getTime()
    };

    var uploaded = Drive.Files.insert(
      resource,
      blob,
      {
        convert: true,
        ocr: true,
        ocrLanguage: 'vi'
      }
    );

    tempDocId = uploaded.id;

    var text = DocumentApp.openById(tempDocId).getBody().getText();

    try {
      DriveApp.getFileById(tempDocId).setTrashed(true);
    } catch (cleanup1) {
      Logger.log('ocr_cleanup_error=' + cleanup1.message);
    }

    return text || '';
  } catch (err) {
    Logger.log('ocr_extract_error=' + err.message);

    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (cleanup2) {
        Logger.log('ocr_cleanup_error=' + cleanup2.message);
      }
    }

    return '';
  }
}

function buildInvoiceExtractionPrompt_(pdfText) {
  return [
    'Hãy trích xuất dữ liệu từ nội dung OCR của hóa đơn VAT tiếng Việt.',
    'Chỉ trả về JSON hợp lệ.',
    'Không giải thích.',
    'Không markdown.',
    'Nếu không chắc chắn thì để chuỗi rỗng.',
    '',
    'Cấu trúc JSON bắt buộc:',
    '{',
    '  "so_hoa_don": "",',
    '  "ky_hieu": "",',
    '  "ngay_hoa_don": "",',
    '  "mst_ben_ban": "",',
    '  "nguoi_ban": "",',
    '  "tien_truoc_thue": "",',
    '  "tien_thue_gtgt": "",',
    '  "tong_tien": "",',
    '  "items": [',
    '    {',
    '      "ten_hang_hoa": "",',
    '      "don_vi_tinh": "",',
    '      "so_luong": "",',
    '      "don_gia": "",',
    '      "thanh_tien": "",',
    '      "thue_suat": "",',
    '      "tien_thue": "",',
    '      "ghi_chu": ""',
    '    }',
    '  ]',
    '}',
    '',
    'Quy tắc:',
    '- Ưu tiên lấy đúng thông tin bên bán.',
    '- Giữ nguyên số hóa đơn, ký hiệu, mã số thuế, tiền.',
    '- Không tự suy diễn.',
    '- Nếu không có bảng hàng hóa thì để items là [].',
    '',
    'Nội dung OCR:',
    pdfText
  ].join('\n');
}

function callOpenAIForInvoiceJson_(apiKey, model, prompt) {
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify({
      model: model,
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [
        {
          role: 'system',
          content: 'Bạn là bộ trích xuất dữ liệu hóa đơn. Chỉ trả về JSON hợp lệ.'
        },
        {
          role: 'user',
          content: prompt
        }
      ]
    })
  });

  var statusCode = response.getResponseCode();
  var body = response.getContentText();

  Logger.log('openai_status=' + statusCode);

  if (statusCode < 200 || statusCode >= 300) {
    Logger.log('openai_error_body=' + body);
    throw new Error('OpenAI API lỗi, mã: ' + statusCode);
  }

  var json = JSON.parse(body);
  var content = (((json || {}).choices || [])[0] || {}).message || {};

  return content.content || '';
}

function parseInvoiceJsonSafe_(rawText) {
  if (!rawText) return null;

  try {
    return JSON.parse(rawText);
  } catch (err1) {
    try {
      var cleaned = String(rawText)
        .replace(/```json/gi, '')
        .replace(/```/g, '')
        .trim();

      var start = cleaned.indexOf('{');
      var end = cleaned.lastIndexOf('}');

      if (start !== -1 && end !== -1 && end > start) {
        cleaned = cleaned.substring(start, end + 1);
      }

      return JSON.parse(cleaned);
    } catch (err2) {
      Logger.log('json_parse_error=' + err2.message);
      return null;
    }
  }
}

function normalizeInvoiceResult_(data) {
  var result = buildEmptyInvoiceResult_();

  result.so_hoa_don = cleanTextValue_(data.so_hoa_don);
  result.ky_hieu = cleanTextValue_(data.ky_hieu);
  result.ngay_hoa_don = cleanTextValue_(data.ngay_hoa_don);
  result.mst_ben_ban = cleanTextValue_(data.mst_ben_ban);
  result.nguoi_ban = cleanTextValue_(data.nguoi_ban);
  result.tien_truoc_thue = cleanTextValue_(data.tien_truoc_thue);
  result.tien_thue_gtgt = cleanTextValue_(data.tien_thue_gtgt);
  result.tong_tien = cleanTextValue_(data.tong_tien);

  if (data.items && Array.isArray(data.items)) {
    result.items = data.items.map(function(item) {
      return {
        ten_hang_hoa: cleanTextValue_(item.ten_hang_hoa),
        don_vi_tinh: cleanTextValue_(item.don_vi_tinh),
        so_luong: cleanTextValue_(item.so_luong),
        don_gia: cleanTextValue_(item.don_gia),
        thanh_tien: cleanTextValue_(item.thanh_tien),
        thue_suat: cleanTextValue_(item.thue_suat),
        tien_thue: cleanTextValue_(item.tien_thue),
        ghi_chu: cleanTextValue_(item.ghi_chu)
      };
    }).filter(function(item) {
      return item.ten_hang_hoa || item.so_luong || item.thanh_tien;
    });
  }

  return result;
}

function buildEmptyInvoiceResult_() {
  return {
    so_hoa_don: '',
    ky_hieu: '',
    ngay_hoa_don: '',
    mst_ben_ban: '',
    nguoi_ban: '',
    tien_truoc_thue: '',
    tien_thue_gtgt: '',
    tong_tien: '',
    items: []
  };
}

function cleanTextValue_(value) {
  return safeString_(value).trim();
}
