const concatTags = require('./concatTags')
const list = require('./list')
const image = require('./image')
const table = require('./table')
const link = require('./link')
const style = require('./style')
// const form = require('./form')

module.exports = (files) => {
  concatTags(files)
  list(files)
  image(files)
  table(files)
  link(files)
  style(files)
  // skip for now https://github.com/jsreport/jsreport/issues/628
  // form(files)
}
