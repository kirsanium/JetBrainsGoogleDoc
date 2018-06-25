function doGet( e ) {
  QUnit.urlParams( e.parameter );
  QUnit.config({ title: "Unit tests for my project" });
  QUnit.load( tests );
  return QUnit.getHtml();
};

QUnit.helpers(this);

function tests() {
  module("tables editing");

  test("thin spaces test", 1, function() {
    ok(testInsertThinSpaces());
  });
}

function prepareTables() {
  var doc = DocumentApp.getActiveDocument();
  var rangeBuilder = doc.newRange();
  var tables = doc.getBody().getTables();
  
  for (var i = 0; i < tables.length; i++) {
    rangeBuilder.addElement(tables[i]);
  }
  doc.setSelection(rangeBuilder.build());
  editSelection();
  
  return tables;
}

function testInsertThinSpaces() {
  var tables = prepareTables();
  var texts = [];
  
  for (var i = 0; i < tables.length; i++) {
    var tableText = getTextChildren(tables[i]);
    texts = texts.concat(tableText);
  }
  
  for (var i = 0; i < texts.length; i++) {
    if (!checkNumber(texts[i].getText())) return false;
  }
    
  return true;
}

function checkNumber(number) {
  var counter = 0;
  var isSpace = false;
  
  for (var i = 0; i < number.length; i++) {
    if (isNumber(number[i]))
      counter++;
    else { 
      if (counter == 3) {
        if (i + 1 < number.length) {
          if (isNumber(number[i+1]) && number[i] != '\u2009') return false;
        }
      }
      counter = 0;
    }
  }
      return true;     
}

