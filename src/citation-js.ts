function doCitationJS(
  templateName: string,
  bibtexes: string[],
  citations: string[][]
) {
  const citationJS = HtmlService.createTemplateFromFile(
    "src/citation-js.lib.html"
  ).getRawContent();
  const citeJS = new Function(
    citationJS + "\nconst cite = require('citation-js'); return cite;"
  )();

  let config = citeJS.plugins.config.get('@csl')
  if (!["apa", "harvard", "vancouver"].includes(templateName)) {
    const template = UrlFetchApp.fetch(
      `https://zotero.org/styles/${templateName}`
    ).getContentText();
    config.templates.add(templateName, template)
  }

  const cite = new citeJS.Cite(bibtexes);
  const bibliography = cite.format("bibliography", {
    format: "text",
    template: templateName,
    lang: "en-US",
  });

  const citationsList = [];
  for (let i = 0; i < citations.length; i++) {
    var citationsPre = citations.slice(0, i).flat();
    var citationsPost = citations.slice(i + 1, citations.length).flat();
    // Remove duplicate keep first
    citationsPost = citationsPost.filter(
      (item, index) => citationsPost.indexOf(item) === index
    );
    var citation = cite.format("citation", {
      format: "text",
      template: templateName,
      lang: "en-US",
      entry: citations[i],
      citationsPre: citationsPre,
      citationsPost: citationsPost,
    });
    citationsList.push(citation);
  }
  return { bibliography: bibliography, citations: citationsList };
}
