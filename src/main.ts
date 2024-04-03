"use strict";

function onOpen(e: GoogleAppsScript.Events.DocsOnOpen) {
  var ui = DocumentApp.getUi();
  ui.createMenu("Bibs for Mendeley")
    .addItem("Connect with Mendeley", "mendeleyLogin")
    .addItem("Setting", "mendeleySetting")
    .addSeparator()
    .addItem("Open library", "openLibrary")
    .addItem("Insert bibliography", "insertBibliography")
    .addToUi();
}

function mendeleyLogin() {
  var service = getService_();
  var authorizationUrl = service.getAuthorizationUrl();
  var template = HtmlService.createTemplate(
    '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
      "Reopen the sidebar when the authorization is complete."
  );
  template.authorizationUrl = authorizationUrl;
  var page = template.evaluate();
  DocumentApp.getUi().showSidebar(page);
}

function mendeleySetting() {
  var template = HtmlService.createTemplateFromFile("templates/setting.html");
  var page = template.evaluate();
  DocumentApp.getUi().showSidebar(page);
}

function saveSetting(citationStyle: string) {
  const ui = DocumentApp.getUi();
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty("citationStyle", citationStyle);
  ui.alert("Setting saved");
}

function openLibrary() {
  var service = getService_();
  var response = UrlFetchApp.fetch(
    "https://api.mendeley.com/folders?maxResults=10",
    {
      headers: {
        Authorization: "Bearer " + service.getAccessToken(),
      },
    }
  );
  var folders = JSON.parse(response.getContentText());

  var template = HtmlService.createTemplateFromFile("templates/libraries.html");
  template.folders = folders;
  var page = template.evaluate();
  DocumentApp.getUi().showSidebar(page);
}

function getDocuments(folder_id: string) {
  var service = getService_();
  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/documents?folder_id=${folder_id}`,
    {
      headers: {
        Authorization: "Bearer " + service.getAccessToken(),
      },
    }
  );
  var documents = JSON.parse(response.getContentText());
  return documents;
}

function getDocumentBibtex(document_id: string) {
  var service = getService_();
  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/documents/${document_id}`,
    {
      headers: {
        Accept: "application/x-bibtex",
        Authorization: "Bearer " + service.getAccessToken(),
      },
      contentType: "application/x-bibtex",
    }
  );
  return response.getContentText();
}

function doCite(documentIDs: string[]) {
  const citationStyle =
    PropertiesService.getDocumentProperties().getProperty("citationStyle");

  var bibtexes = [];
  for (var i = 0; i < documentIDs.length; i++) {
    var documentID = documentIDs[i];
    var bibtex = getDocumentBibtex(documentID);
    var cache = CacheService.getScriptCache();
    cache.put(`bibtex-${documentID}`, bibtex);
    bibtexes.push(bibtex);
  }

  var apiConvert = UrlFetchApp.fetch(
    `https://bibtex-converter.wikidepia.workers.dev/convert/${citationStyle}`,
    {
      method: "post",
      payload: JSON.stringify(bibtexes),
    }
  );
  var apiResult = JSON.parse(apiConvert.getContentText());

  // Insert citation
  insertCitation(apiResult["citation"], documentIDs);

  // Insert bibliography
  insertBibliography(false);
}

function insertNewBibliographyIndex(bibliography: string, documentID: string) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Check if table exists
  var tables = body.getTables();
  var table = null as GoogleAppsScript.Document.Table | null;
  for (var i = 0; i < tables.length; i++) {
    var tableText = tables[i].getCell(0, 0).editAsText().getLinkUrl(0) || "";
    if (tableText.includes("#bibs-mendeley")) {
      table = tables[i];
    }
  }

  if (!table) {
    return;
  }
  // Append new row
  var row = table.appendTableRow();
  var cell = row.appendTableCell();
  cell.setText(bibliography.trim());

  cell
    .editAsText()
    .insertText(0, `​`) // Zero-width space
    .setLinkUrl(0, 0, `#bibs-mendeley-${documentID}`);
  cell.setPaddingBottom(0).setPaddingTop(0);
  cell.setPaddingLeft(0).setPaddingRight(0);
}

// From: https://stackoverflow.com/a/59711275
function insertCitation(citation: string, documentIDs: string[]) {
  var baseDoc = DocumentApp.getActiveDocument();

  var cursor = baseDoc.getCursor();
  if (!cursor) {
    return;
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
    `#cite-mendeley+${documentIDs.join('|')}+${citation.length}`
  );
}

function insertBibliography(createNew: boolean = true) {
  var baseDoc = DocumentApp.getActiveDocument();
  var body = baseDoc.getBody();

  var cites = [];
  var citesSearch = body.findText(`​`);
  while (citesSearch != null) {
    var element = citesSearch.getElement();
    var link = element.asText().getLinkUrl(citesSearch.getStartOffset());
    if (link && link.includes("#cite-mendeley")) {
      var documentIDs = link.split("#cite-mendeley+")[1].split("+")[0].split("|");
      for (var i = 0; i < documentIDs.length; i++) {
        cites.push(documentIDs[i]);
      }
    }
    citesSearch = body.findText(`​`, citesSearch);
  }

  // Check if table already exists
  var tables = body.getTables();
  var table = null as GoogleAppsScript.Document.Table | null;
  for (var i = 0; i < tables.length; i++) {
    var tableText = tables[i].getCell(0, 0).editAsText().getLinkUrl(0) || "";
    if (tableText.includes("#bibs-mendeley")) {
      table = tables[i];
    }
  }

  // Create new table if clicked from menu bar
  if (!table && createNew) {
    var cursor = baseDoc.getCursor();
    if (!cursor) {
      return;
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

  var bibtexes = [];
  var cache = CacheService.getScriptCache();
  for (var i = 0; i < cites.length; i++) {
    var documentID = cites[i];
    var bibtex = cache.get(`bibtex-${documentID}`);
    if (!bibtex) {
      bibtex = getDocumentBibtex(documentID);
      cache.put(`bibtex-${documentID}`, bibtex);
    }
    bibtexes.push(bibtex);
  }

  const citationStyle =
    PropertiesService.getDocumentProperties().getProperty("citationStyle");
  var apiConvert = UrlFetchApp.fetch(
    `https://bibtex-converter.wikidepia.workers.dev/convert/${citationStyle}`,
    {
      method: "post",
      payload: JSON.stringify(bibtexes),
    }
  );
  var apiResult = JSON.parse(apiConvert.getContentText());
  var biblios = apiResult["bibliography"].split("\n");

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
}
