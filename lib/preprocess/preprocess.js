const concatTags = require('./concatTags')
const list = require('./list')
const image = require('./image')
const table = require('./table')
const link = require('./link')

module.exports = (files) => {
  concatTags(files)
  list(files)
  image(files)
  table(files)
  link(files)
}
