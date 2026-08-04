//----------------------------------
// Panel Builder Frontend
//----------------------------------

function servePanelBuilderApp_() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Panel Builder")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1"
    );

}

//----------------------------------
// Include HTML Partial
//----------------------------------

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}
