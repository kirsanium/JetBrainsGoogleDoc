function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Test');
  DocumentApp.getUi().showSidebar(ui);
}

function insertThinSpaces(e) {
  //var regExp = "(\\s)|(\\A)\\d{4,}(\\s)|(\\z)"
  var regExp = "\\d{4,}"
  var textToReplace = e.findText(regExp);
  while (textToReplace) {
    var forReplacement = textToReplace.getElement().asText().editAsText();
    for (var i = textToReplace.getEndOffsetInclusive() - 2; i > textToReplace.getStartOffset(); i = i - 3) {
      forReplacement.insertText(i, " ");
    }
    textToReplace = e.findText(regExp, textToReplace);
  }
  return e;
}

function insertHorizontalRuleToTable(table, rowNum) {
  var row = table.getRow(rowNum);
  var num = row.getNumCells();
  var newRow = table.insertTableRow(rowNum + 1);
  for (var j = 0; j < num; j++) {
    var cell = newRow.appendTableCell();
    cell.setPaddingBottom(0);
    cell.setPaddingLeft(0);
    cell.setPaddingRight(0);
    cell.setPaddingTop(0);
    if (j == 0) cell.insertHorizontalRule(0);
    cell.editAsText().setText("").setFontSize(6);
  }
  return table;
}

function setStyleguideAlignment(e) {
  var regexWord = "[^0-9\\.\\s\\-]";
  var regexNum = "(\\s)*(\\d)+(\\.(\\d)+)?(\\s)*";
  var regexInterval = regexNum + "\\-" + regexNum;
  if (e.getText().search(regexWord) != -1) {
    e.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  }
  else
  {
    var aligned = false;
    var found = e.getText().match(regexNum);
    if (found) {
      for (var i = 0; i < found.length; i++) {
        if (found[i] == e.getText()) {
          e.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
          aligned = true;
          break;
        }
      }
    }
    if (!aligned) {
      found = e.getText().match(regexInterval);
      if (found) {
        for (var i = 0; i < found.length; i++) {
          if (found[i] == e.getText()) {
            e.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            break;
          }
        }
      }
    }
  }
  return e;
}

function formatTables() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var tables = body.getTables();
  for (var i = 0; i < tables.length; i++) {
    tables[i].setBorderWidth(0);
    if (!tables[i].getCell(1, 0).getChild(0).asParagraph()
        .findElement(DocumentApp.ElementType.HORIZONTAL_RULE))
    insertHorizontalRuleToTable(tables[i], 0);
    var rowNum = tables[i].getNumRows();
    for (var j = 0; j < rowNum; j++) {
      var row = tables[i].getRow(j);
      var cellNum = row.getNumCells();
      for (var k = 0; k < cellNum; k++) {
        var cell = row.getCell(k);
        setStyleguideAlignment(cell);
        insertThinSpaces(cell);
        
      }
    }
  }
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
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


