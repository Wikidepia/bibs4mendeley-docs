"use strict";

// https://stackoverflow.com/a/3561711
function escapeRegex(str: string) {
  return str.replace(/[/\-\\^$*+?.()|[\]{}]/g, "\\$&");
}

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
  var cache = CacheService.getUserCache();
  var documents = cache.get(`documents-${folder_id}`);
  if (!documents) {
    var service = getService_();
    var response = UrlFetchApp.fetch(
      `https://api.mendeley.com/documents?folder_id=${folder_id}`,
      {
        headers: {
          Authorization: "Bearer " + service.getAccessToken(),
        },
      }
    );
    documents = response.getContentText();
    cache.put(`documents-${folder_id}`, documents, 24 * 60 * 60);
  }
  return JSON.parse(documents);
}

function getDocumentBibtex(document_id: string) {
  var cache = CacheService.getUserCache();
  var bibtex = cache.get(`bibtex-${document_id}`);
  if (bibtex) {
    return bibtex;
  }

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
  bibtex = response.getContentText();
  cache.put(`bibtex-${document_id}`, bibtex, 24 * 60 * 60);
  return bibtex;
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

  // Insert temp citation
  insertCitation("(temp)", documentIDs);

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
    `#cite-mendeley+${documentIDs.join("|")}+${citation.length}`
  );
}

function insertBibliography(createNew: boolean = true) {
  var baseDoc = DocumentApp.getActiveDocument();
  var body = baseDoc.getBody();

  var bibtexIDs = [];
  var cites = [];
  var citesSearch = body.findText(`​`);
  while (citesSearch != null) {
    var element = citesSearch.getElement();
    var link = element.asText().getLinkUrl(citesSearch.getStartOffset());
    if (link && link.includes("#cite-mendeley")) {
      var xx = [];
      var documentIDs = link
        .split("#cite-mendeley+")[1]
        .split("+")[0]
        .split("|");
      for (var i = 0; i < documentIDs.length; i++) {
        var bibtex = getDocumentBibtex(documentIDs[i]);
        var bibtexID = bibtex.split("\n")[0].split("{")[1].split(",")[0];
        xx.push(bibtexID);
        cites.push(documentIDs[i]);
      }
      bibtexIDs.push(xx);
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

  var bibtexes = cites.map((documentID) => getDocumentBibtex(documentID));
  const citationStyle =
    PropertiesService.getDocumentProperties().getProperty("citationStyle");
  var apiConvert = UrlFetchApp.fetch(
    `https://bibtex-converter.wikidepia.workers.dev/convert/${citationStyle}`,
    {
      method: "post",
      payload: JSON.stringify({ bibtexes: bibtexes, citations: bibtexIDs }),
    }
  );
  var apiResult = JSON.parse(apiConvert.getContentText());
  var biblios = apiResult["bibliography"].split("\n");

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
  var citeMarkOffsets = [];
  var citeCnt = 0;
  var citesSearch = body.findText(`​`);
  while (citesSearch != null) {
    var newCitesSearch = body.findText(`​`, citesSearch);
    var element = citesSearch.getElement();
    var link = element.asText().getLinkUrl(citesSearch.getStartOffset());
    if (link && link.includes("#cite-mendeley")) {
      citeMarkOffsets.push(citesSearch.getStartOffset());
      var citation = apiResult["citations"][citeCnt];
      var curCiteLength = parseInt(link.split("+")[2]);
      var ciText = element.asText();
      ciText.deleteText(
        citesSearch.getStartOffset() - curCiteLength,
        citesSearch.getStartOffset()
      );
      ciText.insertText(
        citesSearch.getStartOffset() - curCiteLength,
        citation + `​`
      );
      var newMarkerOffset =
        citesSearch.getStartOffset() - curCiteLength + citation.length;
      ciText.setLinkUrl(
        newMarkerOffset,
        newMarkerOffset,
        link.slice(0, -1) + citation.length.toString()
      );
      citeCnt++;
    }
    citesSearch = newCitesSearch;
  }
}
