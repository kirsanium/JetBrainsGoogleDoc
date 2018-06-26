/**
 * @OnlyCurrentDoc
**/

// Shows sidebar.
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Edit document');
  DocumentApp.getUi().showSidebar(ui);
}

function onInstall(e) {
  setDefaultUserProperties();
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Show sidebar', 'showSidebar')
      .addItem('Properties', 'showMenu')
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
  var attrs = table.getAttributes();
  for (var att in attrs) {
   Logger.log(att + ":" + attrs[att]);
 }
  table.setBorderWidth(0);
  var rowNum = table.getNumRows();
  var ruleFound = false;
  
  for (var i = 0; i < rowNum; ++i) {
    var row = table.getRow(i);
    var cellNum = row.getNumCells();
    for (var j = 0; j < cellNum; ++j) {
      if (table.getCell(i, j).getChild(0)
          .findElement(DocumentApp.ElementType.HORIZONTAL_RULE)) {
        ruleFound = true;
        break;
      }
            /****TESTING****/
      if (i == 0 && j == 0) 
      {        attrs = table.getCell(i, j).getAttributes(); 

        for (var att in attrs) {
   Logger.log(att + ":" + attrs[att]);
 }
//       // attrs[DocumentApp.Attribute.BORDER_COLOR] = '#FF0000';
//        attrs[DocumentApp.Attribute.BORDER_WIDTH] += 3;
//        table.getCell(i, j).setAttributes(attrs);
      }
      /****TESTING****/
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
  var point = parseFloat(PropertiesService.getUserProperties().getProperty('listIndentation'));
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

function changeStyle() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = getParagraphChildren(body);
  var userProperties = PropertiesService.getUserProperties();
  var tableCellFontSize = 10;
  var normalAttrs = {};
  var titleAttrs = {};
  var header1Attrs = {};
  var header2Attrs = {};
  var header3Attrs = {};
  normalAttrs[DocumentApp.Attribute.FONT_FAMILY]  =
  titleAttrs[DocumentApp.Attribute.FONT_FAMILY]   =
  header1Attrs[DocumentApp.Attribute.FONT_FAMILY] = 
  header2Attrs[DocumentApp.Attribute.FONT_FAMILY] = 
  header3Attrs[DocumentApp.Attribute.FONT_FAMILY] = "Open Sans";
  
  normalAttrs[DocumentApp.Attribute.LINE_SPACING]  = parseFloat(userProperties.getProperty('normalLineSpacing'));
  header1Attrs[DocumentApp.Attribute.LINE_SPACING] = parseFloat(userProperties.getProperty('header1LineSpacing'));
  header2Attrs[DocumentApp.Attribute.LINE_SPACING] = parseFloat(userProperties.getProperty('header2LineSpacing')); 
  header3Attrs[DocumentApp.Attribute.LINE_SPACING] = parseFloat(userProperties.getProperty('header3LineSpacing'));
  titleAttrs[DocumentApp.Attribute.LINE_SPACING]   = parseFloat(userProperties.getProperty('titleLineSpacing'));
  
  normalAttrs[DocumentApp.Attribute.SPACING_BEFORE]  = parseFloat(userProperties.getProperty('normalSpacingBefore'));
  header1Attrs[DocumentApp.Attribute.SPACING_BEFORE] = parseFloat(userProperties.getProperty('header1SpacingBefore'));
  header2Attrs[DocumentApp.Attribute.SPACING_BEFORE] = parseFloat(userProperties.getProperty('header2SpacingBefore'));
  header3Attrs[DocumentApp.Attribute.SPACING_BEFORE] = parseFloat(userProperties.getProperty('header3SpacingBefore'));
  titleAttrs[DocumentApp.Attribute.SPACING_BEFORE]   = parseFloat(userProperties.getProperty('titleSpacingBefore'));
  
  normalAttrs[DocumentApp.Attribute.SPACING_AFTER]  = parseFloat(userProperties.getProperty('normalSpacingAfter'));
  header1Attrs[DocumentApp.Attribute.SPACING_AFTER] = parseFloat(userProperties.getProperty('header1SpacingAfter'));
  header2Attrs[DocumentApp.Attribute.SPACING_AFTER] = parseFloat(userProperties.getProperty('header2SpacingAfter'));
  header3Attrs[DocumentApp.Attribute.SPACING_AFTER] = parseFloat(userProperties.getProperty('header3SpacingAfter'));
  titleAttrs[DocumentApp.Attribute.SPACING_AFTER]   = parseFloat(userProperties.getProperty('titleSpacingAfter'));
  
  normalAttrs[DocumentApp.Attribute.FONT_SIZE]  = parseInt(userProperties.getProperty('normalFontSize'));
  header1Attrs[DocumentApp.Attribute.FONT_SIZE] = parseInt(userProperties.getProperty('header1FontSize'));
  header2Attrs[DocumentApp.Attribute.FONT_SIZE] = parseInt(userProperties.getProperty('header2FontSize'));
  header3Attrs[DocumentApp.Attribute.FONT_SIZE] = parseInt(userProperties.getProperty('header3FontSize'));
  titleAttrs[DocumentApp.Attribute.FONT_SIZE]   = parseInt(userProperties.getProperty('titleFontSize'));
  
  normalAttrs[DocumentApp.Attribute.BOLD]  = isTrue(userProperties.getProperty('normalBold'));
  header1Attrs[DocumentApp.Attribute.BOLD] = isTrue(userProperties.getProperty('header1Bold'));
  header2Attrs[DocumentApp.Attribute.BOLD] = isTrue(userProperties.getProperty('header2Bold'));
  header3Attrs[DocumentApp.Attribute.BOLD] = isTrue(userProperties.getProperty('header3Bold'));
  titleAttrs[DocumentApp.Attribute.BOLD]   = isTrue(userProperties.getProperty('titleBold')); 
  
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL,   normalAttrs);
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.TITLE,    titleAttrs);
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, header1Attrs);
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, header2Attrs);
  body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, header3Attrs);

  for (var i = 0; i < paragraphs.length; i++) {
    var headingType = paragraphs[i].getHeading();
    
    switch(headingType) {
      case DocumentApp.ParagraphHeading.NORMAL:
        paragraphs[i].setAttributes(normalAttrs);
        break;
      case DocumentApp.ParagraphHeading.HEADING1:
        paragraphs[i].setAttributes(header1Attrs);
        break;
      case DocumentApp.ParagraphHeading.HEADING2:
        paragraphs[i].setAttributes(header2Attrs);
        break;
      case DocumentApp.ParagraphHeading.HEADING3:
        paragraphs[i].setAttributes(header3Attrs);
        break;
      case DocumentApp.ParagraphHeading.TITLE:
        paragraphs[i].setAttributes(titleAttrs);
        break;
    }
    if (paragraphs[i].getParent().getType() == DocumentApp.ElementType.TABLE_CELL)
      paragraphs[i].editAsText().setFontSize(tableCellFontSize);
  }
}


function getParagraphChildren(element) {
  var elements = [];
  
  for (var i = 0; i < element.getNumChildren(); i++) {
    var child = element.getChild(i);
    
    if (child.getType() == DocumentApp.ElementType.PARAGRAPH)
      elements.push(child);
    
    if (child.getNumChildren)
      elements = elements.concat(getParagraphChildren(child));
  }
  
  return elements;
}

function swapRows(table, firstRowIndex, secondRowIndex) {
  var secondRow = table.removeRow(secondRowIndex);
  var firstRow = table.removeRow(firstRowIndex);
  table.insertTableRow(firstRowIndex, secondRow);
  table.insertTableRow(secondRowIndex, firstRow);
}

function swapAdjacentRows() {
  var cursor = DocumentApp.getActiveDocument().getCursor();
    
  if (!cursor) 
    throw 'Please, click on row';
  var tableRow = cursor.getElement().getParent();
  
  while (tableRow.getType() != DocumentApp.ElementType.TABLE_ROW) {
    tableRow = tableRow.getParent(); 
    
    if (!tableRow)
      throw 'Please, click on row';
  }
  var table = tableRow.getParentTable();
  var index = table.getChildIndex(tableRow);
  
  if (index < table.getNumRows() - 1)
    swapRows(table, index, index + 1)
}

function isLarger(firstString, secondString) {
  firstString = firstString.replace(/\u2009/g, '');
  secondString = secondString.replace(/\u2009/g, '');
  var firstNumber = parseInt(firstString);
  var secondNumber = parseInt(secondString);
  
  if (firstNumber == NaN || secondNumber == NaN)
    return firstString > secondString;
  else 
    return firstNumber > secondNumber;
}

function sortTable(isDesc) {
  var cursor = DocumentApp.getActiveDocument().getCursor();
    
  if (!cursor) 
    throw 'Please, click on column';
  var tableCell = cursor.getElement().getParent();
  
  while (tableCell.getType() != DocumentApp.ElementType.TABLE_CELL) {
    tableCell = tableCell.getParent();
  }
  var tableRow = tableCell.getParent();
  var colIndex = tableRow.getChildIndex(tableCell);
  var table = tableRow.getParentTable();
  var size = table.getNumRows();
  
  for (var i = 1; i < size; i++) {
    for (var j = size - 1; j > i; j--) {
      var firstCell = table.getRow(j - 1).getCell(colIndex);
      var secondCell = table.getRow(j).getCell(colIndex);
      var isFirstLarger = isLarger(firstCell.getText(), secondCell.getText());
      
      if (isFirstLarger && isDesc)
        swapRows(table, j - 1, j);
      else if (!isFirstLarger && !isDesc)
        swapRows(table, j - 1, j);
    }
  }
}

//Sets default styleguide values.
function setDefaultUserProperties() {
    PropertiesService.getUserProperties().setProperties({
      'listIndentation'      : '7.2',
    
      'normalLineSpacing'    : '1.3',
      'normalSpacingBefore'  : '0',
      'normalSpacingAfter'   : '0',
      'normalFontSize'       : '11',
      'normalBold'           : 'false',
    
      'titleLineSpacing'     : '1.2',
      'titleSpacingBefore'   : '50',
      'titleSpacingAfter'    : '20',
      'titleFontSize'        : '30',
      'titleBold'            : 'true',
    
      'header1LineSpacing'   : '1.3',
      'header1SpacingBefore' : '0',
      'header1SpacingAfter'  : '8',
      'header1FontSize'      : '20',
      'header1Bold'          : 'true',
    
      'header2LineSpacing'   : '1.3',
      'header2SpacingBefore' : '0',
      'header2SpacingAfter'  : '8',
      'header2FontSize'      : '14',
      'header2Bold'          : 'true',
    
      'header3LineSpacing'   : '1.3',
      'header3SpacingBefore' : '0',
      'header3SpacingAfter'  : '2',
      'header3FontSize'      : '11',
      'header3Bold'          : 'true',
  });
}

function applyProperties(properties) {
  var lowerLineSpacing = 1;
  var upperLineSpacing = 100;
  var lowerBASpacing = 0;
  var upperBASpacing = 250;
  var lowerFontSize = 6;
  var upperFontSize = 400;
  
  var error = 'Line spacing: '+lowerLineSpacing.toString()+'-'+upperLineSpacing.toString()+', spacing before/after: '+lowerBASpacing.toString()+
              '-'+upperBASpacing.toString()+', font size: '+lowerFontSize.toString()+'-'+upperFontSize.toString()+'(must be integer)';
  
  var normalLineSpacing    = parseFloat(properties['normalLineSpacing']);
  var normalSpacingBefore  = parseFloat(properties['normalSpacingBefore']);
  var normalSpacingAfter   = parseFloat(properties['normalSpacingAfter']);
  var normalFontSize       = parseFloat(properties['normalFontSize']);
  
  var titleLineSpacing     = parseFloat(properties['titleLineSpacing']);
  var titleSpacingBefore   = parseFloat(properties['titleSpacingBefore']);
  var titleSpacingAfter    = parseFloat(properties['titleSpacingAfter']);
  var titleFontSize        = parseFloat(properties['titleFontSize']);
  
  var header1LineSpacing   = parseFloat(properties['header1LineSpacing']);
  var header1SpacingBefore = parseFloat(properties['header1SpacingBefore']);
  var header1SpacingAfter  = parseFloat(properties['header1SpacingAfter']);
  var header1FontSize      = parseFloat(properties['header1FontSize']);
  
  var header2LineSpacing   = parseFloat(properties['header2LineSpacing']);
  var header2SpacingBefore = parseFloat(properties['header2SpacingBefore']);
  var header2SpacingAfter  = parseFloat(properties['header2SpacingAfter']);
  var header2FontSize      = parseFloat(properties['header2FontSize']);
  
  var header3LineSpacing   = parseFloat(properties['header3LineSpacing']);
  var header3SpacingBefore = parseFloat(properties['header3SpacingBefore']);
  var header3SpacingAfter  = parseFloat(properties['header3SpacingAfter']);
  var header3FontSize      = parseFloat(properties['header3FontSize']);
  
  var listIndentation      = parseFloat(properties['listIndentation']);
  
  if ((properties['normalLineSpacing']  !== normalLineSpacing.toString() )||(normalLineSpacing  < lowerLineSpacing)||(normalLineSpacing  > upperLineSpacing)||
      (properties['titleLineSpacing']   !== titleLineSpacing.toString()  )||(titleLineSpacing   < lowerLineSpacing)||(titleLineSpacing   > upperLineSpacing)||
      (properties['header1LineSpacing'] !== header1LineSpacing.toString())||(header1LineSpacing < lowerLineSpacing)||(header1LineSpacing > upperLineSpacing)||
      (properties['header2LineSpacing'] !== header2LineSpacing.toString())||(header2LineSpacing < lowerLineSpacing)||(header3LineSpacing > upperLineSpacing)||
      (properties['header3LineSpacing'] !== header3LineSpacing.toString())||(header3LineSpacing < lowerLineSpacing)||(header3LineSpacing > upperLineSpacing)||
      
      (properties['normalSpacingBefore']  !== normalSpacingBefore.toString() )||(normalSpacingBefore  < lowerBASpacing)||(normalSpacingBefore  > upperBASpacing)||
      (properties['titleSpacingBefore']   !== titleSpacingBefore.toString()  )||(titleSpacingBefore   < lowerBASpacing)||(titleSpacingBefore   > upperBASpacing)||
      (properties['header1SpacingBefore'] !== header1SpacingBefore.toString())||(header1SpacingBefore < lowerBASpacing)||(header1SpacingBefore > upperBASpacing)||
      (properties['header2SpacingBefore'] !== header2SpacingBefore.toString())||(header2SpacingBefore < lowerBASpacing)||(header2SpacingBefore > upperBASpacing)||
      (properties['header3SpacingBefore'] !== header3SpacingBefore.toString())||(header3SpacingBefore < lowerBASpacing)||(header3SpacingBefore > upperBASpacing)||
        
      (properties['normalSpacingAfter']  !== normalSpacingAfter.toString() )||(normalSpacingAfter  < lowerBASpacing)||(normalSpacingAfter  > upperBASpacing)||
      (properties['titleSpacingAfter']   !== titleSpacingAfter.toString()  )||(titleSpacingAfter   < lowerBASpacing)||(titleSpacingAfter   > upperBASpacing)||
      (properties['header1SpacingAfter'] !== header1SpacingAfter.toString())||(header1SpacingAfter < lowerBASpacing)||(header1SpacingAfter > upperBASpacing)||
      (properties['header2SpacingAfter'] !== header2SpacingAfter.toString())||(header2SpacingAfter < lowerBASpacing)||(header2SpacingAfter > upperBASpacing)||
      (properties['header3SpacingAfter'] !== header3SpacingAfter.toString())||(header3SpacingAfter < lowerBASpacing)||(header3SpacingAfter > upperBASpacing)||
        
      (properties['normalFontSize']  !== normalFontSize.toString() )||(normalFontSize  % 1 !== 0)||(normalFontSize  < lowerFontSize)||(normalFontSize  > upperFontSize)||
      (properties['titleFontSize']   !== titleFontSize.toString()  )||(titleFontSize   % 1 !== 0)||(titleFontSize   < lowerFontSize)||(titleFontSize   > upperFontSize)||
      (properties['header1FontSize'] !== header1FontSize.toString())||(header1FontSize % 1 !== 0)||(header1FontSize < lowerFontSize)||(header1FontSize > upperFontSize)||
      (properties['header2FontSize'] !== header2FontSize.toString())||(header2FontSize % 1 !== 0)||(header2FontSize < lowerFontSize)||(header2FontSize > upperFontSize)||
      (properties['header3FontSize'] !== header3FontSize.toString())||(header3FontSize % 1 !== 0)||(header3FontSize < lowerFontSize)||(header3FontSize > upperFontSize)||
        
      (properties['listIndentation'] !== listIndentation.toString())||(listIndentation < 0))
      {
        throw error;
      }
          
    PropertiesService.getUserProperties().setProperties(properties);
}


function showMenu() {
  var menu = HtmlService.createHtmlOutputFromFile('PropertiesDialog')
      .setWidth(600)
      .setHeight(400);
  DocumentApp.getUi()
      .showModalDialog(menu, 'Properties');
}
 
function getUserProperties() {
  var userProperties = PropertiesService.getUserProperties();
  return [ userProperties.getProperty('normalLineSpacing'),
           userProperties.getProperty('normalSpacingBefore'),
           userProperties.getProperty('normalSpacingAfter'),
           userProperties.getProperty('normalFontSize'),
           userProperties.getProperty('normalBold'),
         
           userProperties.getProperty('titleLineSpacing'),
           userProperties.getProperty('titleSpacingBefore'),
           userProperties.getProperty('titleSpacingAfter'),
           userProperties.getProperty('titleFontSize'),
           userProperties.getProperty('titleBold'),
         
           userProperties.getProperty('header1LineSpacing'),
           userProperties.getProperty('header1SpacingBefore'),
           userProperties.getProperty('header1SpacingAfter'),
           userProperties.getProperty('header1FontSize'),
           userProperties.getProperty('header1Bold'),
         
           userProperties.getProperty('header2LineSpacing'),
           userProperties.getProperty('header2SpacingBefore'),
           userProperties.getProperty('header2SpacingAfter'),
           userProperties.getProperty('header2FontSize'),
           userProperties.getProperty('header2Bold'),
         
           userProperties.getProperty('header3LineSpacing'),
           userProperties.getProperty('header3SpacingBefore'),
           userProperties.getProperty('header3SpacingAfter'),
           userProperties.getProperty('header3FontSize'),
           userProperties.getProperty('header3Bold'),
         
           userProperties.getProperty('listIndentation')];
}

function isTrue(s) {
  if (s === 'true')
    return true;
  else return false;
}