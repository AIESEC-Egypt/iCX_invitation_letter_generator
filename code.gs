function dataExtraction(query) {
  query = JSON.stringify({ query: query });
  var requestOptions = {
    method: "post",
    payload: query,
    contentType: "application/json",
    headers: {
      access_token: "",
    },
  };
  var response = UrlFetchApp.fetch(
    `https://gis-api.aiesec.org/graphql?access_token=${requestOptions["headers"]["access_token"]}`,
    requestOptions
  );
  var recievedDate = JSON.parse(response.getContentText());
  return recievedDate.data.allOpportunityApplication.data;
}

function approvalsUpdating() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs");
  var startDate = "01/11/2021";
  var query = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n\t\t\topportunity_home_mc:1609\n\t\tdate_approved:{from:\"${startDate}\"}\n\t\t}\n  \n\tpage:1\n    per_page:4000\n\t  \n\t)\n\t{\n  \t\n\t\tdata{\n\t\t\tperson{\n\t\t\t\tid\n\t\t\t\tfull_name\n\t\t\t\t\n\t\t\t\temail\n\t\t\t\thome_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n\t\t\t}\n\t\t\topportunity{\n\t\t\t\tid\n\t\t\t}\n\t\t\tdate_approved\n\t\t\tslot{\n\t\t\t\tstart_date\n\t\t\t\tend_date\n\t\t\t}\n\t\t\thost_lc{\n\t\tid\n\t\tname\n\t\t\t}\n\t\t\t\n\t\t\t\n\t\t}\n\t\t\n  }\n}`;
  var data = dataExtraction(query);
  var rows = [];
  for (let i = 0; i < data.length; i++) {
    var searchingID = data[i].person.id + "_" + data[i].opportunity.id;
    var rowIndex = sheet
      .createTextFinder(`${searchingID}`)
      .matchEntireCell(true)
      .findAll()
      .map((x) => x.getRow());
    if (rowIndex.length == 0) {
      GmailApp.sendEmail(
        `${data[i].person.email}`,
        "Invitation Letter of AIESEC in Egypt Form",
        "Greetings from AIESEC in Egypt!\n\nIn this mail you will find an invitation letter form for your internship. In case of any questions, please contact local committee.\nInvitation Link Form: https://docs.google.com/forms/d/e/1FAIpQLSe28BzFSfZng53hMcnjBNsbgxNo9sH5sTqUfUVpfvUw9Dd4Jg/viewform \n\nPlease do not reply to this mail, it's automatical."
      );

      rows.push([
        data[i].person.id + "_" + data[i].opportunity.id,
        data[i].person.id,
        data[i].opportunity.id,
        data[i].person.full_name,
        data[i].person.email,
        data[i].date_approved != null
          ? data[i].date_approved.toString().substring(0, 10)
          : "-",
        data[i].slot.start_date,
        data[i].slot.end_date,
        data[i].person.home_lc.name,
        data[i].person.home_mc.name,
        data[i].host_lc.name,
      ]);
    }
  }
  if (rows.length > 0)
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
}

function sendIL() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RESPONSES");
  const approvalSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs");
  const referenceSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference");
  var data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][22] != "TRUE" && data[i][1] != "") {
      try {
        var row = approvalSheet
          .createTextFinder(`${data[i][18]}`)
          .matchEntireCell(true)
          .findAll()
          .map((x) => x.getRow());
        if (row.length > 0) {
          var title = template
            .getName()
            .toString()
            .replace("{{full_name}}", `${data[i][1]}`);
          var invitationLetter = DocumentApp.create(title);
          Logger.log(invitationLetter.getUrl());
          sheet.getRange(i + 2, 23).setValue(true);
          var invitationLetterBody = invitationLetter.getBody();
          var total_items = template.getBody().getNumChildren();
          for (let i = 0; i < total_items; i++) {
            switch (template.getBody().getChild(i).getType()) {
              case DocumentApp.ElementType.PARAGRAPH:
                invitationLetterBody.appendParagraph(
                  template.getBody().getChild(i).copy()
                );
                break;
              case DocumentApp.ElementType.INLINE_IMAGE:
                invitationLetterBody.appendImage(
                  template.getBody().getChild(i).copy()
                );
                break;
            }
          }
          DriveApp.getFileById(invitationLetter.getId()).moveTo(folder);
          sheet.getRange(i + 2, 24).setValue(invitationLetter.getId());
          var today = Utilities.formatDate(new Date(), "GMT+2", "dd-MM-yyyy");
          invitationLetterBody.replaceText("{{date_of_creation}}", today);
          var numOfIL = referenceSheet
            .getRange(
              referenceSheet
                .createTextFinder(data[i][19])
                .matchEntireCell(true)
                .findAll()
                .map((x) => x.getRow()),
              2
            )
            .getValue();
          var id =
            String(lc_codes[`${data[i][19]}`]) +
            String(numOfIL + 1).padStart(4, "0");
          invitationLetterBody.replaceText("{{IL_ID}}", id);
          referenceSheet
            .getRange(
              referenceSheet
                .createTextFinder(data[i][19])
                .matchEntireCell(true)
                .findAll()
                .map((x) => x.getRow()),
              2
            )
            .setValue(numOfIL + 1);
          sheet.getRange(i + 2, 21).setValue(id);
          invitationLetterBody.replaceText("{{full_name}}", data[i][1]);
          invitationLetterBody.replaceText("{{b_day}}", data[i][3]);
          invitationLetterBody.replaceText("{{place_of_birth}}", data[i][6]);
          invitationLetterBody.replaceText("{{eng_citizenship}}", data[i][4]);
          invitationLetterBody.replaceText("{{passport_id}}", data[i][9]);
          invitationLetterBody.replaceText("{{date_of_issue}}", data[i][7]);
          invitationLetterBody.replaceText("{{expire_date}}", data[i][8]);
          invitationLetterBody.replaceText("{{living_address}}", data[i][5]);
          invitationLetterBody.replaceText("{{project_city}}", data[i][13]);
          invitationLetterBody.replaceText(
            "{{project_start_date}}",
            data[i][14]
          );
          invitationLetterBody.replaceText("{{project_end_date}}", data[i][15]);
          invitationLetter.saveAndClose();
          var lcFolder = DriveApp.getFolderById(`${lcFolders[data[i][19]]}`);
          var file = DriveApp.getFileById(invitationLetter.getId());
          var docblob = file.getBlob().getAs("application/pdf");
          var newDoc = DriveApp.createFile(docblob);
          newDoc.setName(title);
          newDoc.moveTo(lcFolder);
          newDoc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
          sheet.getRange(i + 2, 22).setValue(newDoc.getUrl());
          sheet.getRange(i + 2, 25).setValue(true);
          var message = `Hello ${
            data[i][1]
          },\nGreeting from AIESEC in Egypt!\n\nIn this mail you will find invitation letter for your internship. In case of any questions, please contact local committee. Make sure that your data in the letter is correct, it's your responsibility.\nYou can also download it to print it: ${newDoc.getUrl()}.\n\n\nPlease do not reply to this mail, it's automatical.`;
          MailApp.sendEmail(
            `${data[i][11]}`,
            "Invitation Letter from AIESEC in Egypt",
            message
          );
          sheet.getRange(i + 2, 26).setValue(true);
        }
      } catch (e) {
        Logger.log(e.toString());
        if (e.toString().includes("Invalid email")) {
          var message = `Hello ${
            data[i][1]
          },\nGreeting from AIESEC in Egypt!\n\nIn this mail you will find invitation letter for your internship. In case of any questions, please contact local committee. Make sure that your data in the letter is correct, it's your responsibility.\nYou can also download it to print it: ${newDoc.getUrl()}.\n\n\nPlease do not reply to this mail, it's automatical.`;
          MailApp.sendEmail(
            `${data[i][17]}`,
            "Invitation Letter from AIESEC in Egypt",
            message
          );
          sheet.getRange(i + 2, 26).setValue(true);
        }
      }
    }
  }
  Logger.log("done");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu().addItem("Send IL", "sendIL").addToUi();
}
