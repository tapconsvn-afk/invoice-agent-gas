function setupProject() {
  setDefaultScriptProperties_();
  ensureMainSheets_();
  ensureMainLabels_();
  Logger.log('setup_done');
}

function testOpenAIConnection() {
  var apiKey = getScriptProperty_('OPENAI_API_KEY', '');
  if (!apiKey) {
    Logger.log('openai_api_key_missing');
    return;
  }

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify({
      model: getScriptProperty_('OPENAI_MODEL', 'gpt-4o-mini'),
      temperature: 0,
      messages: [
        {
          role: 'user',
          content: 'Trả lời đúng 1 từ: OK'
        }
      ]
    }),
    muteHttpExceptions: true
  });

  Logger.log('openai_test_status=' + response.getResponseCode());
  Logger.log(response.getContentText());
}
