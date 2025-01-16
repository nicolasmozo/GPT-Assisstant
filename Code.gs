API_KEY = "XXXXXXXXXX";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Sidebar')
    .addItem('Open GPT Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('GPT Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

function fetchGPTResponses(payload) {
  const apiKey = API_KEY;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const { model, inputRange, outputRange } = payload;

  try {
    const inputs = sheet.getRange(inputRange).getValues();
    const results = [];

    inputs.forEach(([input]) => {
      if (input) {
        const url = 'https://api.openai.com/v1/chat/completions';
        const requestPayload = {
          model: model,
          messages: [
            { role: 'system', content: 'You are a helpful assistant.' },
            { role: 'user', content: input }
          ],
          max_tokens: 100,
          temperature: 0.7
        };

        const options = {
          method: 'post',
          contentType: 'application/json',
          headers: {
            Authorization: `Bearer ${apiKey}`
          },
          payload: JSON.stringify(requestPayload),
          muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        Logger.log(response.getContentText());

        const json = JSON.parse(response.getContentText());
        const reply = json.choices[0]?.message?.content?.trim() || 'No response received';
        results.push([reply]);
      } else {
        results.push(['No input provided']);
      }
    });

    sheet.getRange(outputRange).setValues(results);

    return { success: true, message: 'Responses written successfully' };

  } catch (error) {
    return { success: false, message: `Error: ${error.message}` };
  }
}