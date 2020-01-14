const style = require('./style')
const image = require('./image')
const link = require('./link')
// const form = require('./form')
const pageBreak = require('./pageBreak')

module.exports = (files) => {
  style(files)
  image(files)
  link(files)
  // form(files)
  pageBreak(files)
}
