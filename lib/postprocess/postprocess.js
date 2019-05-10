const { DOMParser, XMLSerializer } = require('xmldom')
const removeTagsPlaceholders = require('./removeTagsPlaceholders')
const list = require('./list')

module.exports = (content) => {
  const doc = new DOMParser().parseFromString(content)
  list(doc)
  removeTagsPlaceholders(doc)
  const res = new XMLSerializer().serializeToString(doc)
  return res
}
