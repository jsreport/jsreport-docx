const concatTags = require('./concatTags')
const list = require('./list')
const table = require('./table')
const link = require('./link')
const style = require('./style')

module.exports = (files) => {
  concatTags(files)
  list(files)
  table(files)
  link(files)
  style(files)
}
