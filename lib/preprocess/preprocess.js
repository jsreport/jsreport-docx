const concatTags = require('./concatTags')
const list = require('./list')
const table = require('./table')

module.exports = (files) => {
  concatTags(files)
  list(files)
  table(files)
}
