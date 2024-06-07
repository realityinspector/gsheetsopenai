function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AI')
      .addItem('Generate Text', 'generateText')
      .addToUi();
}

function generateText() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var apiKey = 'sk-';  // Replace with your OpenAI API key
  var lastRequestTime = PropertiesService.getScriptProperties().getProperty('lastRequestTime');
  var currentTime = new Date().getTime();
  var minTimeBetweenRequests = 1000;  // Adjust this value based on rate limits

  if (lastRequestTime && (currentTime - lastRequestTime) < minTimeBetweenRequests) {
    SpreadsheetApp.getUi().alert('Please wait before making another request.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('lastRequestTime', currentTime);

  // Get the active cell
  var cell = sheet.getActiveCell();
  var row = cell.getRow();
  var col = cell.getColumn();
  var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  var values = range.getValues()[0];

  // Remove the value of the active cell from the prompt
  values.splice(col - 1, 1);

  // Clean and combine values
  var cleanedValues = values.map(function(value) {
    return value.toString().trim().replace(/[^a-zA-Z0-9\s]/g, '');  // Remove errant characters
  }).filter(function(value) {
    return value.length > 0;  // Remove empty values
  });

  if (cleanedValues.length === 0) {
    return;
  }

  var combinedPrompt = cleanedValues.join(' ');

  var payload = {
    "model": "gpt-4",
    "messages": [
      {"role": "system", "content": "You are a helpful assistant. "},
      {"role": "user", "content": combinedPrompt}
    ],
    "temperature": 0.7,  // You can adjust this value
    "max_tokens": 3000,   // You can adjust this value
    "top_p": 1.0         // You can adjust this value
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    var json = JSON.parse(response.getContentText());
    var output = json.choices[0].message.content.trim();
    cell.setValue(output);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}
