const ss = SpreadsheetApp.openById("1cvf6qRzOhVgqX8a3iQkceyf_thKx6OcudB35g2wdbFI");

function generate_files(e) {
  // Get the Google Form linked to the response
  var responseSheet = ss.getSheetByName("All Customers");
  var googleForm = FormApp.openById("1q-C9CNnFe5ZfuB8x8yAQsp0nHsvytJuOy3ujbyDA1N4");

  // Get the form response based on the timestamp
  var timestamp = responseSheet.getRange(e.range.getRow(), 1).getValue();
  var formResponse = googleForm.getResponses(timestamp).pop();

  var response_id = formResponse.getId();

  var responseSheet = ss.getSheetByName("All Customers");
  var values = responseSheet.getRange(e.range.getRow(), 2, 1, 23).getDisplayValues();

  var parent_folder = DriveApp.getFolderById("13x_Hr8kFsn-FxBdwVcOd-mNmks99qH3Q");
  var form_21 = DriveApp.getFileById("1GJgzn9okxKzJjQBVL-6435IguQk4zdNO0ktWMGFtQDI");
  var form_20 = DriveApp.getFileById("1Ix2ZfO-0VEc9Kh75moFQxBjV2lkZMSdWqjW8j59ODo8");
  var cust_folder;

  if (values[0][19] == 'Yes') {
    if (check_edit(response_id) == -1 || values[0][22] == '') {
      cust_folder = parent_folder.createFolder(values[0][3].concat("(", values[0][0], ")"));
    }
    else {
      cust_folder = DriveApp.getFolderById(values[0][22].replace(/^.+\//, ''));
      cust_folder.setName(values[0][3].concat("(", values[0][0], ")"));
      while (cust_folder.getFiles().hasNext()) {
        const file = cust_folder.getFiles().next();
        file.setTrashed(true);
    }
    }

    var copy_21 = form_21.makeCopy("Form 21", cust_folder);
    var copy_20 = form_20.makeCopy("Form_20", cust_folder);

    var body_21 = DocumentApp.openById(copy_21.getId()).getBody();
    body_21.replaceText("{{ vehicle_model }}", values[0][12]);
    body_21.replaceText("{{ cust_name }}", values[0][3]);
    body_21.replaceText("{{ bill_date }}", values[0][1]);
    body_21.replaceText("{{ s_w_d }}", values[0][4]);
    body_21.replaceText("{{ cust_add }}", values[0][5]);
    body_21.replaceText("{{ hyp }}", values[0][7]);
    body_21.replaceText("{{ maker }}", values[0][14]);
    body_21.replaceText("{{ chassis }}", values[0][15]);
    body_21.replaceText("{{ motor }}", values[0][16]);
    body_21.replaceText("{{ seating }}", values[0][17]);
    body_21.replaceText("{{ color }}", values[0][18]);

    var body_20 = DocumentApp.openById(copy_20.getId()).getBody();
    body_20.replaceText("{{ cust_name }}", values[0][3]);
    body_20.replaceText("{{ s_w_d }}", values[0][4]);
    body_20.replaceText("{{ cust_add }}", values[0][5]);
    body_20.replaceText("{{ veh_stat }}", values[0][13]);
    body_20.replaceText("{{ maker }}", values[0][14]);
    body_20.replaceText("{{ chassis }}", values[0][15]);
    body_20.replaceText("{{ motor }}", values[0][16]);
    body_20.replaceText("{{ seating }}", values[0][17]);
    body_20.replaceText("{{ color }}", values[0][18]);


    responseSheet.getRange(e.range.getRow(), 24).setValue(cust_folder.getUrl());
  }
}

function duplicate_sheet(customer_name) {
  const templatesheet = ss.getSheetByName("Format");
  var new_sheet = ss.insertSheet(customer_name, ss.getSheets().length, {template: templatesheet});
  var extrasheet = ss.getSheetByName('Copy of Format');
  try {
  ss.deleteSheet(extrasheet);
  } catch(er) {
    
  }
  return new_sheet;
}

function rename_sheet(old_name, new_name) {
  var sheet = ss.getSheetByName(old_name).setName(new_name);
  return sheet;
}

function check_edit(response_id) {
  var id_sheet = ss.getSheetByName("Response Ids");
  var response_ids = id_sheet.getDataRange().getValues();
  let index = response_ids.findIndex( row => row[0] == response_id );
  return index;
}

function set_sheet(e) {
  // Get the Google Form linked to the response
  var responseSheet = ss.getSheetByName("All Customers");
  var googleForm = FormApp.openById("1q-C9CNnFe5ZfuB8x8yAQsp0nHsvytJuOy3ujbyDA1N4");

  // Get the form response based on the timestamp
  var timestamp = responseSheet.getRange(e.range.getRow(), 1).getValue();
  var formResponse = googleForm.getResponses(timestamp).pop();

  var id_sheet = ss.getSheetByName("Response Ids");
  var response_id = formResponse.getId();

  var responseSheet = ss.getSheetByName("All Customers");
  var answers = responseSheet.getRange(e.range.getRow(), 2, 1, 23).getValues();
  var bill_no = answers[0][0];
  var bill_date = answers[0][1];
  var cust_name = answers[0][3];
  var s_w_d = answers[0][4];
  var cust_add = answers[0][5];
  if (cust_add.length > 70) {
    var first_str = cust_add.substring(0, 70);
    var second_str = cust_add.substring(70, cust_add.length);
    cust_add = first_str.concat("\n", second_str);
  }


  var cust_ph = answers[0][6];
  var gur_name = answers[0][9];
  var gur_add = answers[0][10];
  if (gur_add.length > 30) {
    var first_str = gur_add.substring(0, 30);
    var second_str = gur_add.substring(30, gur_add.length);
    gur_add = first_str.concat("\n", second_str);
  }
  var gur_ph = answers[0][11];

  var cust_sheet;

  if (check_edit(response_id) == -1) {
    var response_row = id_sheet.getLastRow()+1;
    id_sheet.getRange(response_row, 1).setValue(response_id);

    cust_sheet = duplicate_sheet(cust_name.concat("(", bill_no, ")"));
    var richtext = SpreadsheetApp.newRichTextValue()
            .setText(cust_name.concat("(", bill_no, ")"))
            .setLinkUrl('#gid=' + cust_sheet.getSheetId())
            .build();
    responseSheet.getRange(e.range.getRow(), 23).setRichTextValue(richtext);
  }
  else {
    var name = responseSheet.getRange(e.range.getRow(), 23).getValue();
    cust_sheet = rename_sheet(name, cust_name.concat("(", bill_no, ")"));
    var richtext = SpreadsheetApp.newRichTextValue()
            .setText(cust_name.concat("(", bill_no, ")"))
            .setLinkUrl('#gid=' + cust_sheet.getSheetId())
            .build();
    responseSheet.getRange(e.range.getRow(), 23).setRichTextValue(richtext);
  }

  cust_sheet.getRange("B1").setValue(bill_date);
  cust_sheet.getRange("D1").setValue(bill_no);
  cust_sheet.getRange("A2:D5").setValue(cust_name.concat(" ", s_w_d, "\n", cust_add, "\n", cust_ph));
  cust_sheet.getRange("E2:F5").setValue(gur_name.concat("\n", gur_add, "\n", gur_ph));
}

function set_edit_url (e) {
  // Get the Google Form linked to the response
  var responseSheet = ss.getSheetByName("All Customers");
  var googleForm = FormApp.openById("1q-C9CNnFe5ZfuB8x8yAQsp0nHsvytJuOy3ujbyDA1N4");

  // Get the form response based on the timestamp
  var timestamp = responseSheet.getRange(e.range.getRow(), 1).getValue();
  var formResponse = googleForm.getResponses(timestamp).pop();

  // Get the Form response URL and add it to the Google Spreadsheet
  var responseUrl = formResponse.getEditResponseUrl();
  var row = e.range.getRow();
  var responseColumn = 22; // Column where the response URL is recorded.
  responseSheet.getRange(row, responseColumn).setValue(responseUrl);

}