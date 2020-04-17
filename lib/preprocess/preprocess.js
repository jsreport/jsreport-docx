const concatTags = require('./concatTags')
const image = require('./image')
const list = require('./list')
const chart = require('./chart')
const table = require('./table')
const link = require('./link')
const style = require('./style')
const pageBreak = require('./pageBreak')

module.exports = (files) => {
  concatTags(files)
  image(files)
  list(files)
  chart(files)
  table(files)
  link(files)
  style(files)
  pageBreak(files)
}
