function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Preview', 'showPreview')
      .addItem('Make PDF', 'makePdf')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function previewText() {
  var text = getSelectedText();
  if (text == "") {
    return DocumentApp.getActiveDocument().getBody().getText();
  } else {
    return text;
  }
}

var message_start_script = "Start your script using <a href=\"http://fountain.io/syntax\">Fountain.io</a> syntax.";
var message_start_script_no_link = "Start your script using Fountain.io syntax.";

function showPreview() {
  var bodyText = previewText();
  var parsedHtml = fountain(bodyText).html;
   
  var previewTemplate = HtmlService.createTemplateFromFile('scriptPreviewTemplate');
  
  if (bodyText != "") {
    previewTemplate.script = parsedHtml.script;
    previewTemplate.titlePage = parsedHtml.title_page;
  } else {
    previewTemplate.script = message_start_script;
    previewTemplate.titlePage = "";
  }
  
  
  var preview = previewTemplate.evaluate();
    
  preview.setTitle('Preview');
  DocumentApp.getUi().showSidebar(preview);
}

function makePdf() {
  if (DocumentApp.getActiveDocument().getBody().getText() == "") {
    DocumentApp.getUi().alert(message_start_script_no_link);
    return;
  }
  
  var pdfPromptTemplate = HtmlService.createTemplateFromFile('makePdfPrompt');
  pdfPromptTemplate.defaultFileName = DocumentApp.getActiveDocument().getName() + ".pdf";
  var pdfPrompt = pdfPromptTemplate.evaluate().setWidth(300).setHeight(70);
  
  DocumentApp.getUi().showModalDialog(pdfPrompt, 'Name the output Pdf');
}

function createPdf(fileName) {
  var bodyText = DocumentApp.getActiveDocument().getBody().getText();
  var parsedHtml = fountain(bodyText).html;
  
  var previewTemplate = HtmlService.createTemplateFromFile('scriptPdfTemplate');
  previewTemplate.script = parsedHtml.script;
  previewTemplate.titlePage = parsedHtml.title_page;
  previewTemplate.driveApp = DriveApp;

  
  var preview = previewTemplate.evaluate();
  
  var pdfBlob = preview.getAs('application/pdf').setName(fileName);  
  
  DriveApp.createFile(pdfBlob);
}
