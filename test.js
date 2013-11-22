var XLS = require('./lib')
  , xls = new XLS(true)

console.log(xls.contentType(3))

console.log(xls.rels())

var sheets = ['This is sheet 1', 'Sheet 2', 'Test']
console.log(xls.app(sheets, 'curapps'))

var name = 'Evan'

console.log(xls.core(name))

console.log(xls.workbook(sheets))