const concatTags = require('./concatTags')
const drawingObject = require('./drawingObject')
const list = require('./list')
const table = require('./table')
const link = require('./link')
const style = require('./style')
const pageBreak = require('./pageBreak')
const multiLine = require('./docxMultiLine')

module.exports = (files) => {
  concatTags(files)
  drawingObject(files)
  list(files)
  table(files)
  link(files)
  style(files)
  pageBreak(files)
  multiLine(files)
}
