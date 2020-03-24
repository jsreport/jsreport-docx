const style = require('./style')
const image = require('./image')
const chart = require('./chart')
const link = require('./link')
const form = require('./form')
const pageBreak = require('./pageBreak')

module.exports = async files => {
  style(files)
  await image(files)
  await chart(files)
  link(files)
  form(files)
  pageBreak(files)
}
