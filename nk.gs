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

function isNumber(character) {
  if (character >= '0' && character <= '9')
    return true;
  
  return false;
}

function insertThinSpaces(text, start, end) {
  if (!start) start = 0;
  if (!end) end = text.getText().length - 1;
  
  var oldText = text.getText().slice(start, end + 1);
  var newText = oldText;
  var numberCount = 0;
  
  for (var i = newText.length - 1; i > 0; i--) {
    
    if (isNumber(newText[i])) {
      numberCount++;
      
      if (numberCount % 3 == 0 && numberCount != 0 && isNumber(newText[i - 1])) {
        newText = [newText.slice(0, i), "\u2009", newText.slice(i)].join('');
      }
    } else {
      numberCount = 0;
    }
  }

  text.deleteText(start, end);
  text.insertText(start, newText)
}

function editSelection() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  if (selection) {
    var elements = selection.getRangeElements();
    
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var textElement = element.getElement().asText();
      
      if (element.isPartial()) {
        insertThinSpaces(textElement, element.getStartOffset(), element.getEndOffsetInclusive());
      } else {
        insertThinSpaces(element);
      }
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var tableCell = cursor.getElement().getParent();
    
    while (tableCell.getType() != DocumentApp.ElementType.TABLE_CELL) {
      tableCell = tableCell.getParent(); 
      
      if (!tableCell)
        throw 'Please, select some text'
    }
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
      var textElements = getTextChildren(currentCell);
      textElements.forEach(function(e) {
                   insertThinSpaces(e);
      });
    }
  }
}

function getTextChildren(element) {
  var elements = [];
  
  for (var i = 0; i < element.getNumChildren(); i++) {
    var child = element.getChild(i);
    Logger.log(child.getType());
    
    if (child.getType() == DocumentApp.ElementType.TEXT)
      elements.push(child);
    
    if (child.getNumChildren)
      elements = elements.concat(getTextChildren(child));
  }
  
  return elements;
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
