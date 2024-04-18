"use strict";

// https://stackoverflow.com/a/3561711
function escapeRegex(str: string) {
  return str.replace(/[/\-\\^$*+?.()|[\]{}]/g, "\\$&");
}

function onOpen(e: GoogleAppsScript.Events.DocsOnOpen) {
  var ui = DocumentApp.getUi();
  ui.createMenu("Bibs for Mendeley")
    .addItem("Connect Mendeley", "mendeleyLogin")
    .addItem("Disconnect Mendeley", "mendeleyLogout")
    .addItem("Citation/Bibs Settings", "mendeleySetting")
    .addSeparator()
    .addItem("Open library", "openLibrary")
    .addItem("Insert bibliography", "insertBibliography")
    .addToUi();
}

function mendeleyLogin() {
  var service = getService_();
  // Check if already authorized
  if (service.hasAccess()) {
    return openLibrary();
  }

  var authorizationUrl = service.getAuthorizationUrl();
  var template = HtmlService.createTemplateFromFile("templates/login.html");
  template.authorizationUrl = authorizationUrl;
  var page = template.evaluate();
  page.setTitle("Bibs for Mendeley Authorization");
  DocumentApp.getUi().showSidebar(page);
}

function mendeleyLogout() {
  var service = getService_();
  service.reset();
  DocumentApp.getUi().alert("Disconnected Mendeley");
}

function mendeleySetting() {
  // Load available citation styles
  var styles = HtmlService.createTemplateFromFile(
    "src/data/csl-nodep.json.html"
  );
  var stylesObj = JSON.parse(styles.getRawContent());

  // Construct citation style options HTML
  const citationStyle =
    PropertiesService.getDocumentProperties().getProperty("citationStyle");
  var stylesHTML = "";
  for (var i = 0; i < stylesObj.length; i++) {
    stylesHTML += `<option value="${stylesObj[i].name}"`;
    if (stylesObj[i].name == citationStyle) {
      stylesHTML += ` selected`;
    }
    stylesHTML += `>${stylesObj[i].title}</option>`;
  }

  var template = HtmlService.createTemplateFromFile("templates/setting.html");
  template.stylesHTML = stylesHTML;
  var page = template.evaluate();
  page.setTitle("Bibs for Mendeley Setting");
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
  page.setTitle("Bibs for Mendeley Libraries");
  DocumentApp.getUi().showSidebar(page);
}

// Fetch documents to be showed in libraries sidebar
// See templates/libraries.html
function getDocuments(folder_id: string) {
  var cache = CacheService.getDocumentCache();
  if (!cache) {
    DocumentApp.getUi().alert(
      "Something went wrong when fetching documents. [ERR: Failed to get cache]"
    );
    throw new Error("Failed to get cache");
  }

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
    cache.put(`documents-${folder_id}`, documents);
  }
  return JSON.parse(documents);
}

function getDocumentBibtex(document_id: string): string {
  var cache = CacheService.getDocumentCache();
  if (!cache) {
    DocumentApp.getUi().alert(
      "Something went wrong when fetching documents. [ERR: Failed to get cache]"
    );
    return "";
  }

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
  cache.put(`bibtex-${document_id}`, bibtex);
  return bibtex;
}

function doCite(documentIDs: string[]) {
  // Insert temp citation
  insertCitation("(BibliographyIsNotInserted)", documentIDs);

  // Insert bibliography
  insertBibliography(false);
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
      var documentIDs = link
        .split("#cite-mendeley+")[1]
        .split("+")[0]
        .split("|");
      for (var i = 0; i < documentIDs.length; i++) {
        var bibtex = getDocumentBibtex(documentIDs[i]);
        if (bibtex.length == 0) {
          DocumentApp.getUi().alert(
            "Failed to fetch bibtex for document ID: " + documentIDs[i]
          );
          return;
        }
        var bibtexID = bibtex.split("\n")[0].split("{")[1].split(",")[0];
        bibtexIDLink.push(bibtexID);
        if (!cites.includes(documentIDs[i])) cites.push(documentIDs[i]);
      }
      bibtexIDs.push(bibtexIDLink);
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
      break;
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
    var link = markerLinks[searchCnt];

    // Skip if link is different, which means
    // it is already updated
    if (afterSearchLink != link) {
      citesSearch = body.findText(`​`, citesSearch);
      continue;
    }
    searchCnt++;

    if (link && link.includes("#cite-mendeley")) {
      const citation = citationJSResult["citations"][citeCnt];
      const curCiteLength = parseInt(link.split("+")[2]);
      const ciText = element.asText();

      // Replace old citation with new citation
      ciText.deleteText(oldMarkerOffset - curCiteLength, oldMarkerOffset);
      ciText.insertText(oldMarkerOffset - curCiteLength, citation + `​`);

      // Update marker linkURL
      const newMarkerOffset = oldMarkerOffset - curCiteLength + citation.length;
      ciText.setLinkUrl(
        newMarkerOffset,
        newMarkerOffset,
        link.split("+").slice(0, 2).join("+") + "+" + citation.length.toString()
      );
      citeCnt++;
    }
    citesSearch = body.findText(`​`, citesSearch);
  }
}
