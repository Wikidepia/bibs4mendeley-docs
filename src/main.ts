"use strict";

function onInstall(e: GoogleAppsScript.Events.DocsOnOpen) {
  onOpen(e);
}

function onOpen(e: GoogleAppsScript.Events.DocsOnOpen) {
  var ui = DocumentApp.getUi();
  var menu = ui.createMenu("Bibs for Mendeley");
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    menu.addItem("Start add-ons", "openUnauthorizedSidebar").addToUi();
    return;
  }

  menu
    .addItem("Connect Mendeley", "mendeleyLogin")
    .addItem("Disconnect Mendeley", "mendeleyLogout")
    .addSeparator()
    .addItem("Open Library", "openLibrary")
    .addItem("Insert Bibliography", "insertBibliography")
    .addToUi();

  // Set default properties
  const documentProperties = PropertiesService.getDocumentProperties();
  if (!documentProperties.getProperty("citationStyle")) {
    documentProperties.setProperty("citationStyle", "apa");
  }
  if (!documentProperties.getProperty("groupID")) {
    documentProperties.setProperty("groupID", "");
  }
}

function openUnauthorizedSidebar() {
  var alertSidebar = HtmlService.createHtmlOutput(
    `Please allow the Add-ons to run. If you already see this message, please refresh the page.
    <br><br>
    If you still see this message after refreshing the page, please enable these add-ons for this document. You can follow the following instructions:
    <ol>
    <li>On your computer, open a document, spreadsheet, or presentation.</li>
    <li>Click <b>Extensions</b> and then <b>Add-ons</b> and then <b>Manage add-ons</b>.</li>
    <li>To turn the add-on on or off, next to the add-on, click <b>Options</b> (three bullet) and then <b>Use in this document</b>.</li>
    </ol>`
  );
  alertSidebar.setTitle("Bibs for Mendeley Unauthorized");
  DocumentApp.getUi().showSidebar(alertSidebar);
}

function mendeleyLogin() {
  // Create authorizationURL
  let authorizationUrl = "https://api.mendeley.com/oauth/authorize";
  authorizationUrl += "?response_type=code";
  authorizationUrl += "&client_id=" + "18291";
  authorizationUrl +=
    "&redirect_uri=" + encodeURIComponent("https://bibs4mendeley.pages.dev");
  authorizationUrl += "&scope=all";
  authorizationUrl += "&state=" + "1234567890";

  var template = HtmlService.createTemplateFromFile("templates/login.html");
  template["authorizationUrl"] = authorizationUrl;
  var page = template.evaluate();
  page.setTitle("Bibs for Mendeley Authorization");
  DocumentApp.getUi().showSidebar(page);
}

function mendeleyLogout() {
  resetAccessToken();

  // Close sidebar (https://stackoverflow.com/a/63844458)
  var closeSidebar = HtmlService.createHtmlOutput(
    "<script>google.script.host.close();</script>"
  );
  DocumentApp.getUi().showSidebar(closeSidebar);
  DocumentApp.getUi().alert("Disconnected Mendeley");
}

function mendeleySetting() {
  const documentProperties = PropertiesService.getDocumentProperties();
  // Load available citation styles
  var styles = HtmlService.createTemplateFromFile(
    "src/data/csl-nodep.json.html"
  );
  var stylesObj = JSON.parse(styles.getRawContent());

  const citationStyle = documentProperties.getProperty("citationStyle");
  const groupID = documentProperties.getProperty("groupID");

  // Construct citation style options HTML
  var stylesHTML = stylesObj
    .map((style: { name: string; title: string }) => {
      var isSelected = style.name == citationStyle ? " selected" : "";
      return `<option value="${style.name}"${isSelected}>${style.title}</option>`;
    })
    .join("");

  // Choose group menu
  var groups = getGroups();
  var defaultOption = `<option value="">None</option>`;

  var groupOptions = groups.map((group: { id: string; name: string }) => {
    var isSelected = group.id == groupID ? " selected" : "";
    return `<option value="${group.id}"${isSelected}>${group.name}</option>`;
  });
  var groupsHTML = defaultOption + groupOptions.join("");

  var template = HtmlService.createTemplateFromFile("templates/setting.html");
  template["stylesHTML"] = stylesHTML;
  template["groupsHTML"] = groupsHTML;
  var page = template.evaluate();
  page.setTitle("Bibs for Mendeley Setting");
  DocumentApp.getUi().showSidebar(page);
}

function saveSetting(citationStyle: string, groupID: string) {
  const documentProperties = PropertiesService.getDocumentProperties();

  if (citationStyle != documentProperties.getProperty("citationStyle")) {
    documentProperties.setProperty("citationStyle", citationStyle);
    insertBibliography(false); // Refresh bibliography
  }
  documentProperties.setProperty("groupID", groupID);
}

function openLibrary() {
  const documentProperties = PropertiesService.getDocumentProperties();
  var groupID = documentProperties.getProperty("groupID");
  if (!groupID) {
    groupID = "";
  }

  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/folders?maxResults=50&group_id=${groupID}`,
    {
      headers: {
        Authorization: "Bearer " + getAccessToken(),
      },
      muteHttpExceptions: true,
    }
  );

  var responseCode = response.getResponseCode();
  if (responseCode === 404) {
    documentProperties.setProperty("groupID", "");
    DocumentApp.getUi().alert("Group not found, and has been set to None.");
    openLibrary();
    return;
  } else if (responseCode === 401) {
    DocumentApp.getUi().alert(
      "You are not connected to Mendeley. Please try to connect again."
    );
    mendeleyLogin();
    return;
  }

  var folders = JSON.parse(response.getContentText());
  var template = HtmlService.createTemplateFromFile("templates/libraries.html");
  template["folders"] = folders;
  var page = template.evaluate();
  page.setTitle("Bibs for Mendeley Libraries");
  DocumentApp.getUi().showSidebar(page);
}

// Fetch documents to be showed in libraries sidebar
// See templates/libraries.html
function getDocuments(folder_id: string) {
  const documentProperties = PropertiesService.getDocumentProperties();
  var groupID = documentProperties.getProperty("groupID");
  if (!groupID) {
    groupID = "";
  }

  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/documents?folder_id=${folder_id}&group_id=${groupID}`,
    {
      headers: {
        Authorization: "Bearer " + getAccessToken(),
      },
      muteHttpExceptions: true,
    }
  );

  var responseCode = response.getResponseCode();
  if (responseCode === 404) {
    DocumentApp.getUi().alert(
      "Folder or group not found, try refreshing the library."
    );
    throw new Error("Folder or group not found");
  } else if (responseCode === 401) {
    DocumentApp.getUi().alert(
      "You are not connected to Mendeley. Please try to connect again."
    );
    mendeleyLogin();
    throw new Error("Not connected to Mendeley");
  }

  const documents = response.getContentText();
  return JSON.parse(documents);
}

function getDocumentBibtex(document_id: string): string {
  var documentProperties = PropertiesService.getDocumentProperties();
  var bibtex = documentProperties.getProperty(`bibtex-${document_id}`);
  if (bibtex) {
    return bibtex;
  }

  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/documents/${document_id}`,
    {
      headers: {
        Accept: "application/x-bibtex",
        Authorization: "Bearer " + getAccessToken(),
      },
      contentType: "application/x-bibtex",
      muteHttpExceptions: true,
    }
  );

  var responseCode = response.getResponseCode();
  if (responseCode === 404) {
    return "";
  } else if (responseCode === 401) {
    DocumentApp.getUi().alert(
      "You are not connected to Mendeley. Please try to connect again."
    );
    mendeleyLogin();
    // Raise exception to stop further execution
    throw new Error("Not connected to Mendeley");
  }

  bibtex = response.getContentText();
  documentProperties.setProperty(`bibtex-${document_id}`, bibtex);
  return bibtex;
}

function getGroups() {
  var response = UrlFetchApp.fetch(`https://api.mendeley.com/groups/v2`, {
    headers: {
      Authorization: "Bearer " + getAccessToken(),
    },
    muteHttpExceptions: true,
  });

  var responseCode = response.getResponseCode();
  if (responseCode === 401) {
    DocumentApp.getUi().alert(
      "You are not connected to Mendeley. Please try to connect again."
    );
    mendeleyLogin();
    throw new Error("Not connected to Mendeley");
  }

  var documents = response.getContentText();
  return JSON.parse(documents);
}

function doCite(documentIDs: string[]) {
  var tempText = insertCitation("(BibliographyIsNotInserted)", documentIDs);

  try {
    insertBibliography(false);
  } catch (e) {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    if (cursor) {
      cursor
        .getElement()
        .asText()
        .deleteText(
          cursor.getSurroundingTextOffset(),
          cursor.getSurroundingTextOffset() + tempText.getText().length - 1
        ); // Remove temp text
    }
  }
}

// From: https://stackoverflow.com/a/59711275
function insertCitation(
  citation: string,
  documentIDs: string[]
): GoogleAppsScript.Document.Text {
  var ui = DocumentApp.getUi();
  var baseDoc = DocumentApp.getActiveDocument();

  var cursor = baseDoc.getCursor();
  if (!cursor) {
    ui.alert(
      "Please place the cursor in the place you want to insert citation."
    );
    throw new Error("cursor not found");
  }

  // Append text to after the cursor
  // var cursorTextOffset = cursor.getSurroundingTextOffset();
  // var cursorTextBefore = cursor.getSurroundingText().getText();
  // if (cursorTextBefore[cursorTextOffset - 1] != " ") {
  //   citation = " " + citation;
  // }

  var text = cursor.insertText(citation);

  // Add marker behind with zws
  text.appendText(`​`);
  text.setLinkUrl(
    text.getText().length - 1,
    text.getText().length - 1,
    `#cite-mendeley+${documentIDs.join("|")}+${citation.length}`
  );
  return text.copy();
}

function insertBibliography(createNew: boolean = true) {
  var ui = DocumentApp.getUi();
  var baseDoc = DocumentApp.getActiveDocument();
  var body = baseDoc.getBody();

  var markerLinks = [] as string[];
  var bibtexIDs = [];
  var cites = [] as string[];
  var citesSearch = body.findText(`​`);
  while (citesSearch != null) {
    var element = citesSearch.getElement();
    var link = element.asText().getLinkUrl(citesSearch.getStartOffset());
    markerLinks.push(link);
    if (link && link.includes("#cite-mendeley")) {
      var bibtexIDLink = [];
      var documentIDs =
        link.split("#cite-mendeley+")?.[1]?.split("+")?.[0]?.split("|") || [];

      if (documentIDs.length == 0) {
        ui.alert("Failed to parse document IDs");
        throw new Error("failed to parse document IDs");
      }

      for (var i = 0; i < documentIDs.length; i++) {
        var documentID = documentIDs[i] || "";
        if (documentID.length == 0) {
          continue;
        }
        var bibtex = getDocumentBibtex(documentID);
        if (bibtex.length == 0) {
          ui.alert("Failed to fetch bibtex for document ID: " + documentID);
          throw new Error(
            "failed to fetch bibtex for document ID: " + documentID
          );
        }
        var bibtexID = bibtex.split("\n")?.[0]?.split("{")?.[1]?.split(",")[0];
        bibtexIDLink.push(bibtexID || "");
        if (!cites.includes(documentID)) cites.push(documentID);
      }
      bibtexIDs.push(bibtexIDLink);
    }
    citesSearch = body.findText(`​`, citesSearch);
  }

  // Check if table already exists
  var allTables = body.getTables();
  var table = undefined as GoogleAppsScript.Document.Table | undefined;
  for (var i = 0; i < allTables.length; i++) {
    var tableText =
      allTables[i]?.getCell(0, 0).editAsText().getLinkUrl(0) || "";
    if (tableText.includes("#bibs-mendeley")) {
      table = allTables[i];
      break;
    }
  }

  // Create new table if clicked from menu bar
  if (!table && createNew) {
    var cursor = baseDoc.getCursor();
    if (!cursor) {
      ui.alert(
        "Please place the cursor in the place you want to insert bibliography."
      );
      throw new Error("cursor not found");
    }

    var cursorPos = baseDoc.getBody().getChildIndex(cursor.getElement());
    table = body.insertTable(cursorPos);
    table.setBorderWidth(0);
  } else if (!table) {
    return;
  }

  // Remove all rows
  for (var i = table.getNumRows() - 1; i >= 0; i--) {
    table.removeRow(i);
  }

  var bibtexes = cites.map((documentID) => getDocumentBibtex(documentID));
  const citationStyle =
    PropertiesService.getDocumentProperties().getProperty("citationStyle");
  if (!citationStyle) {
    return;
  }
  const citationJSResult = doCitationJS(citationStyle, bibtexes, bibtexIDs);
  var biblios = citationJSResult["bibliography"].split("\n");

  // Write bibliography to table
  for (var i = 0; i < biblios.length; i++) {
    if (biblios[i].trim().length == 0) {
      continue;
    }
    var row = table.appendTableRow();
    var cell = row.appendTableCell();
    cell.setText(biblios[i].trim());

    var cellText = cell.editAsText();
    cellText
      .insertText(0, `​`) // Zero-width space
      .setLinkUrl(0, 0, `#bibs-mendeley-${cites[i]}`);
    cell.setPaddingBottom(0).setPaddingTop(0);
    cell.setPaddingLeft(0).setPaddingRight(0);
  }

  // Update all citations with new index
  var searchCnt = 0,
    citeCnt = 0;
  var citesSearch = body.findText(`​`);
  while (citesSearch != null) {
    const oldMarkerOffset = citesSearch.getStartOffset();
    var element = citesSearch.getElement();
    var afterSearchLink = element.asText().getLinkUrl(oldMarkerOffset);
    var markerLink = markerLinks[searchCnt];

    // Skip if link is different, which means
    // it is already updated
    if (afterSearchLink != markerLink) {
      citesSearch = body.findText(`​`, citesSearch);
      continue;
    }
    searchCnt++;

    if (markerLink && markerLink.includes("#cite-mendeley")) {
      const citation = citationJSResult["citations"][citeCnt];
      const curCiteLength = parseInt(markerLink?.split("+")?.[2] || "0");
      if (citation.length == 0 || curCiteLength == 0) {
        DocumentApp.getUi().alert(
          "failed to fetch citation for document ID: " + cites[searchCnt]
        );
        throw new Error(
          "failed to fetch citation for document ID: " + cites[searchCnt]
        );
      }

      const ciText = element.asText();
      // Replace old citation with new citation
      ciText.deleteText(oldMarkerOffset - curCiteLength, oldMarkerOffset);
      ciText.insertText(oldMarkerOffset - curCiteLength, citation + `​`);

      // Update marker linkURL
      const newMarkerOffset = oldMarkerOffset - curCiteLength + citation.length;
      ciText.setLinkUrl(
        newMarkerOffset,
        newMarkerOffset,
        markerLink.split("+").slice(0, 2).join("+") +
          "+" +
          citation.length.toString()
      );
      citeCnt++;
    }
    citesSearch = body.findText(`​`, citesSearch);
  }
}
