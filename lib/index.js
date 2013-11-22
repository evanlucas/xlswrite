var npmlog  = require('npmlog')
  , fs      = require('fs')
  , path    = require('path')
  , builder = require('xmlbuilder')
  , moment  = require('moment')

module.exports = XLS

function XLS(debug) {
  this.log = npmlog
  this.debug = debug || false
  this.log.level = debug ? 'verbose' : 'error'
}

/**
 * Files
 *
 *  /[Content-Types].xml
 *  /_rels/.rels
 *  /docProps/app.xml
 *  /docProps/core.xml
 *  /xl/comments.xml
 *  /xl/endnotes.xml
 *  /xl/fontTable.xml
 *  /xl/footer1.xml
 *  /xl/footer2.xml
 *  /xl/footer3.xml
 *  /xl/footer4.xml
 *  /xl/footnotes.xml
 *  /xl/header1.xml
 *  /xl/header2.xml
 *  /xl/header3.xml
 *  /xl/header4.xml
 *  /xl/header5.xml
 *  /xl/header6.xml
 *  /xl/numbering.xml
 *  /xl/settings.xml
 *  /xl/styles.xml
 *  /xl/theme/theme1.xml
 *  
 */

XLS.prototype.rels = function() {
  var item = builder.create('Relationships', {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true
  })
  
  item
    .ele('Relationship')
    .att('Id', 'rId1')
    .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
    .att('Target', 'xl/workbook.xml')
  
  item
    .ele('Relationship')
    .att('Id', 'rId2')
    .att('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties')
    .att('Target', 'docProps/core.xml')
  
  item
    .ele('Relationship')
    .att('Id', 'rId3')
    .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties')
    .att('Target', 'docProps/app.xml')
  
  return item.end({pretty: this.debug})
}

XLS.prototype.contentType = function(sheetCount) {
  var item = builder.create('Types', { 
    xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types',
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true
  })
  item
    .ele('Default')
    .att('Extension', 'xml')
    .att('ContentType', 'application/xml')
  item
    .ele('Default')
    .att('Extension', 'rels')
    .att('ContentType', 'application/vnd.openxmlformats-package.relationships+xml')
  item
    .ele('Default')
    .att('Extension', 'jpeg')
    .att('ContentType', 'image/jpeg')
  item
    .ele('Override')
    .att('PartName', '/xl/workbook.xml')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
  
  for (var i=1; i<= sheetCount; i++) {
    item
      .ele('Override')
      .att('PartName', '/xl/worksheets/sheet'+i+'.xml')
      .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
  }
  
  // TODO
  // Possibly support multiple themes?
  item
    .ele('Override')
    .att('PartName', '/xl/theme/theme1.xml')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml')
  
  item
    .ele('Override')
    .att('PartName', '/xl/styles.xml')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml')
  
  item
    .ele('Override')
    .att('PartName', '/docProps/core.xml')
    .att('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml')
  
  item
    .ele('Override')
    .att('PartName', '/docProps/app.xml')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')
  
  return item.end({pretty: this.debug})
}

XLS.prototype.app = function(sheets, company, vers) {
  var item = builder.create('Properties', {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types',
    'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true
  })
  
  item.ele('Application').txt('Excel')
  item.ele('DocSecurity').txt('0')
  item.ele('ScaleCrop').txt('false')
  item.ele('HeadingPairs')
    .ele('vt:vector')
    .att('size', 2)
    .att('baseType', 'variant')
      .ele('vt:variant')
        .ele('vt:lpdtr', 'Worksheets')
        .up().up()
      .ele('vt:variant')
        .ele('vt:i4').txt(sheets.length)
        .up()
        
  var titles = item.ele('TitlesOfParts')
    .ele('vt:vector')
    .att('size', sheets.length)
    .att('baseType', 'lpstr')
  
  sheets.forEach(function(sheet) {
    titles.ele('vt:lpstr').txt(sheet)
  })
  
  item.ele('Company').txt(company || '')
  item.ele('LinksUpToDate').txt('false')
  item.ele('SharedDoc').txt('false')
  item.ele('HyperlinksChanged').txt('false')
  item.ele('AppVersion').txt(vers || '14.0300')

  return item.end({pretty: this.debug})
}

XLS.prototype.core = function(name) {
  var date = new Date()
  var item = builder.create('cp:coreProperties', {
    'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
    'xmlns:dcterms': 'http://purl.org/dc/terms/',
    'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
    'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true
  })
  
  item.ele('dc:creator', name || '')
  item.ele('cp:lastModifiedBy', name || '')
  item.ele('dcterms:created')
    .att('xsi:type', 'dcterms:W3CDTF')
    .txt(moment(date).toISOString())
  item.ele('dcterms:modified')
    .att('xsi:type', 'dcterms:W3CDTF')
    .txt(moment(date).toISOString())
  
  return item.end({pretty: this.debug})
}

XLS.prototype.workbook = function(sheets) {
  var item = builder.create('workbook', {
    version: '1.0',
    encoding: 'UTF-8',
    standalone: true,
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  })
  
  item.ele('fileVersion')
    .att('appName', 'xl')
    .att('lastEdited', '5')
    .att('lowestEdited', '5') // Calculate
    .att('rupBuild', '20225') // Find out
  
  item.ele('workbookPr')
    .att('showInkAnnotation', '0')
    .att('autoCompressPictures', '0')
  
  item.ele('bookViews')
    .ele('workbookView')
    .att('xWindow', '0')
    .att('yWindow', '0')
    .att('windowWidth', '700')
    .att('windowHeigth', '700')
    .att('tabRatio', '500')
    .att('activeTab', '2')
  
  var sheetsEle = item.ele('sheets')
  sheets.forEach(function(sheet, idx) {
    sheetsEle.ele('sheet')
      .att('name', sheet)
      .att('sheetId', (idx+1))
      .att('r:id', 'rId'+(idx+1))
  })
  
  item.ele('calcPr')
    .att('calcId', '140000')
    .att('concurrentCalc', '0')
  
  item.ele('extLst')
    .ele('ext')
    .att('xmlns:mx', 'http://schemas.microsoft.com/office/mac/excel/2008/main')
    .att('uri', '{7523E5D3-25F3-A5E0-1632-64F254C22452}')
    .ele('mx:ArchID')
    .att('Flags', '2')
  
  return item.end({pretty: this.debug})
}