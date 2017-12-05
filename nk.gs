/**
 * @OnlyCurrentDoc
**/

// Shows sidebar.
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Edit document');
  DocumentApp.getUi().showSidebar(ui);
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function showError() {
  var ui = DocumentApp.getUi();
  
}

// Checks if the character is number.
// @param character - a symbol.
function isNumber(character) {
  if (character >= '0' && character <= '9')
    return true;
  
  return false;
}

// Inserts thins spaces in all numbers in text between start and end indices.
// @param text - text to insert thin spaces into.
// @param start - start index.
// @param end - end index.
function insertThinSpaces(text, start, end) {
  if (!start) start = 0;
  if (!end) end = text.getText().length - 1;
  if (end - start <= 1)
    return;
  var oldText = text.getText().slice(start, end + 1);
  var newText = oldText;
  var numberCount = 0;
  
  newText = newText.replace(/ /g, '');
  
  for (var i = newText.length - 1; i > 0; i--) {
    
    if (isNumber(newText[i])) {
      numberCount++;
      
      if (numberCount % 3 === 0 && numberCount !== 0 && isNumber(newText[i - 1])) {
        newText = [newText.slice(0, i), "\u2009", newText.slice(i)].join('');
      }
    } else {
      numberCount = 0;
    }
  }
  
  text.deleteText(start, end);
  text.insertText(start, newText);
}

// Inserts thin spaces into a selected piece of a document.
function editSelection() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  if (selection) {
    var elements = selection.getRangeElements();
    
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var textElement = element.getElement();
      
      if (textElement.getType() != DocumentApp.ElementType.TEXT) {
        var textElements = getTextChildren(textElement);
        textElements.forEach(function(e) {
          insertThinSpaces(e);
        });
        
        continue;
      }
      
      if (element.isPartial()) {
        insertThinSpaces(textElement, element.getStartOffset(), element.getEndOffsetInclusive());
      } else {
        insertThinSpaces(textElement);
      }
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    
    if (!cursor) 
      throw 'Please, select some text';
    var tableCell = cursor.getElement().getParent();
    
    while (tableCell.getType() != DocumentApp.ElementType.TABLE_CELL) {
      tableCell = tableCell.getParent(); 
      
      if (!tableCell)
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
    
    if (child.getType() == DocumentApp.ElementType.TEXT)
      elements.push(child);
    
    if (child.getNumChildren)
      elements = elements.concat(getTextChildren(child));
  }
  
  return elements;
}

// Inserts horizontal rule to table after row #rowNum.
// @param table - table to insert the rule into.
// @param rowNum - number of a row to insert the rule after.
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
    if (j === 0) cell.insertHorizontalRule(0);
    cell.editAsText().setText("").setFontSize(6);
  }
  
  return table;
}

// Sets alignment according to the styleguide.
// @param e - element to apply the alignment to.
function setStyleguideAlignment(table, column) {
  var regexWord = /[^0-9\.\s\- %]/;
  var regexNum = /\s*[\d ]+.(?=[\d ]+)?[\d ]*%?\s*/;
  var regexInterval = new RegExp(regexNum.source + "\-" + regexNum.source);
  var rowNum = table.getNumRows();
  var cell = table.getCell(rowNum - 1, column);
  
  while (!(/\S/.test(cell.getText())) && rowNum > 1) {
    rowNum -= 1;
    cell = table.getCell(rowNum - 1, column);
  }
  
  if (cell.getText().search(regexWord) != -1) {
    
    for (var i = 0; i < rowNum; ++i) {
      var cell2align = table.getCell(i, column);
      if (cell2align.getColSpan() == 1) {
        var numChild = cell2align.getNumChildren();
        for (var j = 0; j < numChild; ++j) {
          cell2align.getChild(j).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        }    
      }
    }
  }
  else {
    var aligned = false;
    var found = cell.getText().match(regexNum);
    
    if (found) {
      
      if (found[0] == cell.getText()) {
        
        for (var i = 0; i < rowNum; ++i) {
          var cell2align = table.getCell(i, column);
          if (cell2align.getColSpan() == 1) {
            var numChild = cell2align.getNumChildren();
            for (var j = 0; j < numChild; ++j) {
              cell2align.getChild(j).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
            }
          }        
        }
        aligned = true;
      }
    }
    if (!aligned) {
      found = cell.getText().match(regexInterval);
      
      if (found) {
        
        if (found[0] == cell.getText()) {
          
          for (var i = 0; i < rowNum; ++i) {
            var cell2align = table.getCell(i, column);
            
            if (cell2align.getColSpan() == 1) {
              var numChild = cell2align.getNumChildren();
              
              for (var j = 0; j < numChild; ++j) {
                cell2align.getChild(j).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
              }
            }
          }
        }
      }
    }
  }
  
  return cell;
}

// Sets alignment according to the styleguide to all selected tables.
function formatSelectedTables() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var isTable = true;
  
  if (selection) {
    var elements = selection.getRangeElements();
    
    for (var i = 0; i < elements.length; i++) {
      isTable = true;      
      var element = elements[i].getElement();
      
      if (element.getType() == DocumentApp.ElementType.TABLE) {
        var numRows = element.getNumRows();
        for (var j = 0; j < numRows; j++) {
          var row = element.getRow(j);
          var numCells = row.getNumCells();
          for (var k = 0; k < numRows; k++) {
            var cell = row.getCell(k);
            var numChildren = cell.getNumChildren();
            for (var f = 0; f < numChildren; f++)
              if (cell.getChild(f).getType() == DocumentApp.ElementType.TABLE)
                formatTable(cell.getChild(f));
          }
        }
      }
      
      if (element.getType() == DocumentApp.ElementType.TABLE_CELL) {
        var numChildren = element.getNumChildren();
        for (var j = 0; j < numChildren; j++)
          if (element.getChild(j).getType() == DocumentApp.ElementType.TABLE)
            formatTable(element.getChild(j));
      }
      
      while (element.getType() != DocumentApp.ElementType.TABLE && element.getType() != DocumentApp.ElementType.DOCUMENT)
      {
        element = element.getParent();
        
        if (element.getType() == DocumentApp.ElementType.DOCUMENT)
          isTable = false;
      }
      
      if (isTable) formatTable(element);
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var element = cursor.getElement();
    
    if (element.getType() == DocumentApp.ElementType.DOCUMENT)
          isTable = false;
    
    while (element.getType() != DocumentApp.ElementType.TABLE && element.getType() != DocumentApp.ElementType.DOCUMENT)
      {
        element = element.getParent();
        
        if (element.getType() == DocumentApp.ElementType.DOCUMENT)
          isTable = false;
      }
    
      if (isTable) formatTable(element);
  }
  
  if (!isTable)
    throw 'Please, select a table';
}

// Sets alignment according to the styleguide to the table.
// @param table - the table to apply he alignment to.
function formatTable(table) {
  table.setBorderWidth(0);
  var rowNum = table.getNumRows();
  var ruleFound = false;
  
  for (var i = 0; i < rowNum; ++i) {
    var row = table.getRow(i);
    var cellNum = row.getNumCells();
    for (var j = 0; j < cellNum; ++j) {
      if (table.getCell(i, j).getChild(0).asParagraph()
          .findElement(DocumentApp.ElementType.HORIZONTAL_RULE)) {
        ruleFound = true;
        break;
      }
    }
  }
  
  if (!ruleFound) {
    insertHorizontalRuleToTable(table, 0);
  }
  
  var row = table.getRow(rowNum - 1);
  var cellNum = row.getNumCells();
    
  for (var k = 0; k < cellNum; k++) {
    setStyleguideAlignment(table, k);
  }
}

// Sets alignment according to the styleguide to all the tables in the document.
function formatAllTables() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var tables = body.getTables();
  
  for (var i = 0; i < tables.length; i++) {
    formatTable(tables[i]);
  }
}

// Sets indentation according to the styleguide to the list item.
// @param listItem - the listItem to set the indentation to.
function formatListItem(listItem) {
  var nestingLevel = listItem.getNestingLevel() + 1;
  var point = 7.2;
  Logger.log(listItem.getIndentFirstLine());
  Logger.log(listItem.getIndentStart());
  listItem.setIndentFirstLine(nestingLevel * point);
  listItem.setIndentStart(nestingLevel * point + 18);
}

// Sets indentation according to the styleguide to all the list items.
function formatAllListItems() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var listItems = body.getListItems();
  
  for (var i = 0; i < listItems.length; i++) {
    formatListItem(listItems[i]);
  }
}

// Sets indentation according to the styleguide to the selected list items.
function formatSelectedListItems() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var isListItem = true;
  
  if (selection) {
    var elements = selection.getRangeElements();
    
    for (var i = 0; i < elements.length; i++) {
      isListItem = true;      
      var element = elements[i].getElement();
      
      if (element.getType() == DocumentApp.ElementType.TABLE) {
        var numRows = element.getNumRows();
        for (var j = 0; j < numRows; j++) {
          var row = element.getRow(j);
          var numCells = row.getNumCells();
          for (var k = 0; k < numRows; k++) {
            var cell = row.getCell(k);
            var numChildren = cell.getNumChildren();
            for (var f = 0; f < numChildren; f++)
              if (cell.getChild(f).getType() == DocumentApp.ElementType.LIST_ITEM)
                formatListItem(cell.getChild(f));
          }
        }
      }
      
      if (element.getType() == DocumentApp.ElementType.TABLE_CELL) {
        var numChildren = element.getNumChildren();
        for (var j = 0; j < numChildren; j++)
          if (element.getChild(j).getType() == DocumentApp.ElementType.LIST_ITEM)
            formatListItem(element.getChild(j));
      }
      
      while (element.getType() != DocumentApp.ElementType.LIST_ITEM && element.getType() != DocumentApp.ElementType.DOCUMENT) {
        element = element.getParent();
        
        if (element.getType() == DocumentApp.ElementType.DOCUMENT)
          isListItem = false;
      }
      
      if (isListItem) formatListItem(element);
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var element = cursor.getElement();
    
    if (element.getType() == DocumentApp.ElementType.DOCUMENT)
          isListItem = false;
    
    while (element.getType() != DocumentApp.ElementType.LIST_ITEM && element.getType() != DocumentApp.ElementType.DOCUMENT) {
        element = element.getParent();
        
        if (element.getType() == DocumentApp.ElementType.DOCUMENT)
          isListItem = false;
      }
    
      if (isListItem) formatListItem(element);
  }
  
  if (!isListItem) 
    throw 'Please, select a list';
}

// Formats the document according to the styleguide.
function formatEverything() {
  var doc = DocumentApp.getActiveDocument();
  var rangeBuilder = doc.newRange();
  var paragraphs = doc.getBody().getParagraphs();
  
  for (var i = 0; i < paragraphs.length; i++) {
    rangeBuilder.addElement(paragraphs[i]);
  }
  
  formatAllTables();
  formatAllListItems();
}