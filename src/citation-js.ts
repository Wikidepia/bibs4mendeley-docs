function doCitationJS(
  template: string,
  bibtexes: string[],
  citations: string[][]
) {
  const citationJS = HtmlService.createTemplateFromFile(
    "src/citation-js.v0.7.9.js.html"
  ).getRawContent();
  const Cite = new Function(
    citationJS + "\nconst Cite = require('citation-js'); return Cite;"
  )();

  const cite = new Cite(bibtexes);
  const bibliography = cite.format("bibliography", {
    format: "text",
    template: template,
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
      template: template,
      lang: "en-US",
      entry: citations[i],
      citationsPre: citationsPre,
      citationsPost: citationsPost,
    });
    citationsList.push(citation);
  }
  return { bibliography: bibliography, citations: citationsList };
}
