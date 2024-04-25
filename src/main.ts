"use strict";

// https://stackoverflow.com/a/3561711
function escapeRegex(str: string) {
  return str.replace(/[/\-\\^$*+?.()|[\]{}]/g, "\\$&");
}

function onInstall(e: GoogleAppsScript.Events.DocsOnOpen) {
  onOpen(e);
}

function onOpen(_e: GoogleAppsScript.Events.DocsOnOpen) {
  var ui = DocumentApp.getUi();
  ui.createMenu("Bibs for Mendeley")
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

function mendeleyLogin() {
  var service = getService_();
  // Check if already authorized
  if (service.hasAccess()) {
    return openLibrary();
  }

  var authorizationUrl = service.getAuthorizationUrl();
  var template = HtmlService.createTemplateFromFile("templates/login.html");
  template["authorizationUrl"] = authorizationUrl;
  var page = template.evaluate();
  page.setTitle("Bibs for Mendeley Authorization");
  DocumentApp.getUi().showSidebar(page);
}

function mendeleyLogout() {
  var service = getService_();
  service.reset();

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
    .map(
      (style: { name: string; title: string }) =>
        `<option value="${style.name}"${
          style.name == citationStyle ? " selected" : ""
        }>${style.title}</option>`
    )
    .join(""); // TODO: Make it clearereerer

  // Choose group menu
  var groups = getGroups();
  var groupsHTML = `<option value="">None</option>`;
  groupsHTML += groups
    .map(
      (group: { id: string; name: string }) =>
        `<option value="${group.id}"${group.id == groupID ? " selected" : ""}>
        ${group.name}</option>`
    )
    .join(""); // TODO: Make it clearereerer

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
  const groupID = documentProperties.getProperty("groupID");

  var service = getService_();
  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/folders?maxResults=50&group_id=${groupID}`,
    {
      headers: {
        Authorization: "Bearer " + service.getAccessToken(),
      },
    }
  );
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
  const groupID = documentProperties.getProperty("groupID");

  var service = getService_();
  var response = UrlFetchApp.fetch(
    `https://api.mendeley.com/documents?folder_id=${folder_id}&group_id=${groupID}`,
    {
      headers: {
        Authorization: "Bearer " + service.getAccessToken(),
      },
    }
  );
  const documents = response.getContentText();
  return JSON.parse(documents);
}

function getDocumentBibtex(document_id: string): string {
  var documentProperties = PropertiesService.getDocumentProperties();

  var bibtex = documentProperties.getProperty(`bibtex-${document_id}`);
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
  documentProperties.setProperty(`bibtex-${document_id}`, bibtex);
  return bibtex;
}

function getGroups() {
  var service = getService_();
  var response = UrlFetchApp.fetch(`https://api.mendeley.com/groups/v2`, {
    headers: {
      Authorization: "Bearer " + service.getAccessToken(),
    },
  });
  var documents = response.getContentText();
  return JSON.parse(documents);
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
        return;
      }

      for (var i = 0; i < documentIDs.length; i++) {
        var documentID = documentIDs[i] || "";
        if (documentID.length == 0) {
          continue;
        }
        var bibtex = getDocumentBibtex(documentID);
        if (bibtex.length == 0) {
          ui.alert(
            "Failed to fetch bibtex for document ID: " + documentID
          );
          return;
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
          "Failed to fetch citation for document ID: " + cites[searchCnt]
        );
        return;
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
