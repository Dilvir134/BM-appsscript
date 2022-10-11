function onOpen(e) {
  let ui = SpreadsheetApp.getUi(); 
  ui.createMenu('🤖 Automation Tools')
    .addItem('Move balance to individual sheet', 'SENDREPAIR')
    .addItem('Move Finance to individual sheet', 'SENDFINANCE')
    .addToUi(); 
}; 

function get_repair_values() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  var values = [];
  var new_values = [];
  values = sheet.getRange("A2:F").getValues();
  for (var i=0; i < values.length; i++) {
  values[i].splice(2, 1);
  values[i].splice(2, 1);
  if (values[i][0] == 'BALANCE') { 
    if (values[i][3].toString().toUpperCase() == 'PAID') {
      var notation = "F"
      notation += (i+2).toString();
      var link = sheet.getRange(notation).getRichTextValue().getLinkUrl();
      values[i][3] = link;
    }
    var rangeToBeLinked;
    rangeToBeLinked = sheet.getRange(i+2, 2);
    var sheetlink = '#gid=' + sheet.getSheetId() + '&range=' + 'B' + rangeToBeLinked.getRow();
    values[i][0] = sheet.getName();
    values[i][4] = sheetlink;
    new_values.push(values[i]);
    }
  }
  return new_values;
}

function SENDREPAIR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const repair_sheet = ss.getSheetByName("Repair Balance");
  var repair_values = get_repair_values();
  var final_values = [];
  var numOfRows = 0;
  var row = 0;
  var delete_values = repair_sheet.getRange("A2:A").getDisplayValues();
  var merged = delete_values.reduce(function(prev, next) {
  return prev.concat(next);
  });

  row = merged.indexOf(sheet.getName());
  numOfRows = merged.filter(x => x==sheet.getName()).length;

  if (row != -1) {
    repair_sheet.deleteRows(row+2, numOfRows);
  }

  for (var i=0; i < repair_values.length; i++) {
    final_values = [];
    final_values.push(repair_values[i].slice(1, 3));
    var j = repair_sheet.getLastRow();
    repair_sheet.getRange(j+1, 2, 1, 2).setValues(final_values);
    var sheettext = SpreadsheetApp.newRichTextValue()
          .setText(repair_values[i][0])
          .setLinkUrl(repair_values[i][4])
          .build();
    repair_sheet.getRange(j+1, 1, 1, 1).setRichTextValue(sheettext);
    if (repair_values[i][3] != '') {
      var text = SpreadsheetApp.newRichTextValue()
          .setText("PAID")
          .setLinkUrl(repair_values[i][3])
          .build();
      repair_sheet.getRange(j+1, 4, 1, 1).setRichTextValue(text);
    }
  }
}

function get_finance_values() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  var values = [];
  var new_values = [];
  values = sheet.getRange("A2:F").getValues();
  for (var i=0; i < values.length; i++) {
  values[i].splice(3, 1);
  values[i].splice(3, 1);
  if ((values[i][0]).includes('(Finance)')) {
    if (values[i][3].toString().toUpperCase() == 'PAID') {
      var notation = "F"
      notation += (i+2).toString();
      var link = sheet.getRange(notation).getRichTextValue().getLinkUrl();
      values[i][3] = link;
    }
    var rangeToBeLinked;
    rangeToBeLinked = sheet.getRange(i+2, 2);
    var sheetlink = '#gid=' + sheet.getSheetId() + '&range=' + 'B' + rangeToBeLinked.getRow();
    values[i][5] = values[i][0];
    values[i][0] = sheet.getName();
    values[i][4] = sheetlink; 
    new_values.push(values[i]);
  }
}
  return new_values;
}

function SENDFINANCE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const finance_sheet = ss.getSheetByName("Finance");
  var finance_values = get_finance_values();
  var final_values = [];
  var numOfRows = 0;
  var row = 0;
  var delete_values = finance_sheet.getRange("A2:A").getDisplayValues();
  var merged = delete_values.reduce(function(prev, next) {
  return prev.concat(next);
  });

  row = merged.indexOf(sheet.getName());
  console.log(merged);
  numOfRows = merged.filter(x => x==sheet.getName()).length;

  if (row != -1) {
    finance_sheet.deleteRows(row+2, numOfRows);
  }

  for (var i=0; i < finance_values.length; i++) {
    final_values = [];
    final_values.push(finance_values[i].slice(1, 3));
    var j = finance_sheet.getLastRow();
    finance_sheet.getRange(j+1, 2, 1, 2).setValues(final_values);
    var sheettext = SpreadsheetApp.newRichTextValue()
          .setText(finance_values[i][0])
          .setLinkUrl(finance_values[i][4])
          .build();
    finance_sheet.getRange(j+1, 1, 1, 1).setRichTextValue(sheettext);
    finance_sheet.getRange(j+1, 5, 1, 1).setValue(finance_values[i][5]);
    if (finance_values[i][3] != '') {
      var text = SpreadsheetApp.newRichTextValue()
          .setText("PAID")
          .setLinkUrl(finance_values[i][3])
          .build();
      finance_sheet.getRange(j+1, 4, 1, 1).setRichTextValue(text);
    }
  }
}