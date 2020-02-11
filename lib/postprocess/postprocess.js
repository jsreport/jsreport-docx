const style = require('./style')
const image = require('./image')
const link = require('./link')
const form = require('./form')
const pageBreak = require('./pageBreak')

module.exports = async files => {
  style(files)
  await image(files)
  link(files)
  form(files)
  pageBreak(files)
}
