const concatTags = require('./concatTags')
const image = require('./image')
const list = require('./list')
const table = require('./table')
const link = require('./link')
const style = require('./style')

module.exports = (files) => {
  concatTags(files)
  image(files)
  list(files)
  table(files)
  link(files)
  style(files)
}
