const style = require('./style')
const drawingObject = require('./drawingObject')
const link = require('./link')
const form = require('./form')
const pageBreak = require('./pageBreak')
const removeBlockHelper = require('./removeBlockHelper')

module.exports = async (files, options) => {
  await pageBreak(files)
  style(files)
  await drawingObject(files, options)
  link(files)
  form(files)
  await removeBlockHelper(files)
}
