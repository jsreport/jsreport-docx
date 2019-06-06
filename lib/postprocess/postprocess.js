const removeTagsPlaceholders = require('./removeTagsPlaceholders')
const style = require('./style')
const image = require('./image')

module.exports = (files) => {
  style(files)
  image(files)
  removeTagsPlaceholders(files)
}
