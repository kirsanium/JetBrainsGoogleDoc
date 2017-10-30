function changeBulletToAsterisk() {
  var body = DocumentApp.getActiveDocument().getBody();
  var listItems = body.getListItems();
  var listItemsCopies = [];
  for (var i = 0; i < listItems.length; i++) {
    var text = listItems[i].getText();
    text = "*  " + text + "\n";
    var bodyText = body.editAsText();
    bodyText.insertText(0, text);
  }
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Test');
  DocumentApp.getUi().showSidebar(ui);
}



function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  if (selection) {
    var text = selection.getRangeElements();
    text = insertThinSpaces(text);
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    
    if (cursor) {
    }
  }
}


function insertThinSpaces(text) {
  var numberCount = 0;
  
  for (var i = text.length - 1; i >= 0; i--) {
    if (!isNan(text.charAt(i))) {
      numberCount++;
    }
    
    if (numberCount % 3 == 0) {
      text = [text.slice(0, i), ' ', text.slice(i)].join('');
    }
  }
  
  return text;
}

function testFunction() {
  console.log("test");
}

//function myFunction() {
//  var doc = DocumentApp.getActiveDocument();
//  var body = doc.getBody();    
//  var element = body.getChild(2).asListItem(); 
//  var attrs = element.getAttributes(); 
//  for (var att in attrs) {
//    Logger.log(att + " : " + attrs[att]);
//  }
//}