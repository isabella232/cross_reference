var types = {
  fig: { name: 'figure ', text: 'Figure ', key: 'fig', five: 'figur', descMarker: 'ഛಎ', indexMarker: '☙', namedRange: 'lofTable' },
  tab: { name: 'table ', text: 'Table ', key: 'tab', five: 'table', descMarker: 'ഭഫ', indexMarker: '❧', namedRange: 'lotTable'  },
  equ: { name: 'equation ', text: 'Equation ', key: 'equ', five: 'equat', descMarker: 'ടച', indexMarker: '❦', namedRange: 'loeTable'  },
  fno: { name: 'footnote ', text: 'Footnote ', key: 'fno', five: 'fnote', descMarker: 'തമ', indexMarker: '❥', namedRange: 'lonTable'  }
};

function createLoF() {
  createIndex( 'fig' );
}

function createLoT() {
  createIndex( 'tab' );
}

function createIndex(key) {
  var type = types[key];
  var cursor = getCursorIndex();
  var labSettings = PropertiesService.getDocumentProperties().getProperty( 'cross_' + type.key );
  var labText = labSettings ?  toCap( labSettings.split( '_' )[ 2 ] ) : type.text;

  if ( updateDoc() === 'error' ) return;
  
  var labCount = encodeLabel(key);
  var position = deleteLoF(key) || cursor;
  
  insertDummyLoF( key, labCount, labText, position );
  
  var template = HtmlService.createTemplateFromFile( 'lof' );
  template.props = type;
  var html = template.evaluate();
  html.setWidth( 250 ).setHeight( 90 );
  DocumentApp.getUi().showModalDialog( html, 'Generating list of ' + type.text );
}

function getCursorIndex() {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  if ( !cursor ) return 0;
  
  var element = cursor.getElement();
  
  return element.getParent().getChildIndex(element);
}


function encodeLabel(key) {
  var type = types[key];
  var doc = DocumentApp.getActiveDocument();
  var paragraphs = doc.getBody().getParagraphs();
  var labCount = {};
  labCount[ type.key ] = 0;
  var indexDescs = '';
  
  for ( var i = 0; i < paragraphs.length; i++ ) {
    var text = paragraphs[ i ].editAsText();
    var locs = getCrossLinks( text, 5 );
    var start = locs[ 0 ][ 0 ];
    var url = locs[ 2 ][ 0 ];

    if ( !locs[ 0 ].length ) continue;
    var str = '#' + type.key;
    if ( url.substr( 0, 4 ) === str ) {

      indexDescs += type.descMarker + text.getText().match(/([ ]\d[^\w]*)([^\.]*)/)[2]
    
      text.deleteText( start, start + 1 )
        .insertText( start, type.indexMarker );
      labCount[ type.key ]++;
    }
  }
  
  PropertiesService.getDocumentProperties().setProperty(type.key + '_descs', indexDescs)
  return labCount;
}


function deleteLoF(key) {
  var indexTable = findLoF(key);
  if ( !indexTable ) return;
  
  var index = indexTable.getParent().getChildIndex( indexTable );
  indexTable.removeFromParent();
  
  return index;
}


function findLoF(key) {
  var type = types[key];
  var index = DocumentApp.getActiveDocument().getNamedRanges( type.namedRange )[ 0 ];
  
  return index ? index.getRange().getRangeElements()[ 0 ].getElement().asTable() : null;
}


function insertDummyLoF( key, labCount, labText, position ) {
  var type = types[key];
  var doc = DocumentApp.getActiveDocument();
  var lofCells = [];
  var labText = toCap( labText );
  var placeholder = '...';
  var range = doc.newRange();
  
  doc.getNamedRanges( type.namedRange ).forEach( function( r ) {
    r.remove()
  });
  
  var indexDescs = PropertiesService.getDocumentProperties().getProperty(type.key + '_descs');
  var splitDescs = indexDescs ? indexDescs.split(type.descMarker) : null;
  
  for ( var i = 1; i <= labCount[ type.key ]; i++ ) {
    var indexName = labText + i;
    var indexDesc = splitDescs && splitDescs[i].length ? ': ' + splitDescs[i] : '';
    var row = [ indexName + indexDesc, placeholder ];
    lofCells.push( row );
  }
  
  var indexTable = doc.getBody().insertTable( position, lofCells )
  styleLoF( indexTable );
  
  range.addElement( indexTable );
  doc.addNamedRange(type.namedRange, range.build() )
}


function styleLoF( indexTable ) {
  
  indexTable.setBorderWidth( 0 );
  
  var styleAttributes = {
    'BOLD': null,
    'ITALIC': null,
    'UNDERLINE': null,
    'FONT_SIZE': null
  };
  
  for ( var i = indexTable.getNumRows(); i--; ) {
    var row = indexTable.getRow( i );
    
    indexTable.setAttributes( styleAttributes ).setColumnWidth( 1, 64 );
    row.getCell( 0 ).setPaddingLeft( 0 );
    row.getCell( 1 ).setPaddingRight( 0 )
      .getChild( 0 ).asParagraph().setAlignment( DocumentApp.HorizontalAlignment.RIGHT );
  }
}


function getDocAsPDF() {
   return DocumentApp.getActiveDocument().getBlob().getBytes();
}


function insertLoFNumbers( key, pg_nums ) {
  var type = types[key];  
  var lofTable = findLoF(key);
  var currentRow = 0;
  
  for ( var i = 0; i < pg_nums.length; i++ ) {
    var labCount = pg_nums[ i ];
    if ( !labCount ) continue;

    for ( var j = currentRow; j < lofTable.getNumRows(); j++ ) {
      lofTable.getCell( j, 1 )
        .clear()
        .getChild( 0 ).asParagraph().appendText( i + 1 );
    }
    var currentRow = currentRow + labCount;
  }
}


function restoreLabels( key ) {
  var type = types[key]; 
  var doc = DocumentApp.getActiveDocument();
  var paras = doc.getBody().getParagraphs();
  
  for ( var i = 0; i < paras.length; i++ ) {
    var text = paras[ i ].editAsText();
    var locs = getCrossLinks( text, 5 );
    var starts = locs[ 0 ];
    var urls = locs[ 2 ];
    
    if ( !starts.length ) continue;
    
    for ( var k = starts.length; k--; ) {
      var start = starts[ k ];
      var url = urls[ k ];
      var str = '#' + type.key;
      if (url.substr( 0, 4 ) === str) text.deleteText( start - 1, start );
    }
  }
  
  updateDoc();
}
