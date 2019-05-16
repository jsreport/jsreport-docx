const { DOMParser, XMLSerializer } = require('xmldom')
const removeTagsPlaceholders = require('./removeTagsPlaceholders')

module.exports = (content) => {
  const doc = new DOMParser().parseFromString(content)
  removeTagsPlaceholders(doc)
  const res = new XMLSerializer().serializeToString(doc)
  return res
}
