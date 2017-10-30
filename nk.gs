function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Test');
  DocumentApp.getUi().showSidebar(ui);
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  if (selection) {
    var text = [];
    var elements = selection.getSelectedElements();
    
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement();
        
        if (element.editAsText) {
          var elementText = element.asText().getText();
          text.push(elementText);
        }
      }
    }
    
    if (text.length == 0) {
      return false;
    }
    
    return text;
  } else {
    return false;
  }
}

function isNumber(character) {
  if (character >= '0' && character <= '9')
    return true;
  
  return false;
}

function insertThinSpaces(string) {
  var numberCount = 0;
  
  for (var i = string.length - 1; i > 0; i--) {
    
    if (isNumber(string[i])) {
      numberCount++;
      
      if (numberCount % 3 == 0 && numberCount != 0 && isNumber(string[i - 1])) {
        string = [string.slice(0, i), "\u2009", string.slice(i)].join('');
      }
    } else {
      numberCount = 0;
    }
  }
  
  return string;
}

function insertText(newText) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var elements = selection.getSelectedElements();
  
  for (var i = 0; i < elements.length; i++) {
    if (elements[i].isPartial()) {
      var element = elements[i].getElement().asText();
      var startIndex = elements[i].getStartOffset();
      var endIndex = elements[i].getEndOffsetInclusive();
      var remainingText = element.getText().substring(endIndex + 1);
      
      element.deleteText(startIndex, endIndex);
      element.insertText(startIndex, newText);
    } else {
      var element = elements[i].getElement();
      
      if (element.editAsText) {
        element.asText().setText(newText[i]);
      }
    }
  }
}

function editSelection() {
   var selectedText = getSelectedText();
  if (selectedText) {
    for (var i = 0; i < selectedText.length; i++) {
      selectedText[i] = insertThinSpaces(selectedText[i]); 
    }
    
    insertText(selectedText);
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var tableCell = cursor.getElement().getParent();
    
    if (tableCell.getType() == DocumentApp.ElementType.PARAGRAPH) {
      tableCell = tableCell.getParent();
    }
    
    if (tableCell.getType() != DocumentApp.ElementType.TABLE_CELL) {
      throw 'Please, select some text';
    }
    
    var colIndex = tableCell.getParent().getChildIndex(tableCell);
    var table = tableCell.getParentTable();
    var numRows = table.getNumChildren();
    
    for (var i = 0; i < numRows; i++) {
      var currentCell = table.getChild(i).getChild(colIndex);
      var text = insertThinSpaces(currentCell.editAsText().getText());
      currentCell.setText(text);
    }
  }
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
