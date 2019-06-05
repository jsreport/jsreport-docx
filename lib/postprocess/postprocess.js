const { DOMParser, XMLSerializer } = require('xmldom')
const removeTagsPlaceholders = require('./removeTagsPlaceholders')
const style = require('./style')

module.exports = (content) => {
  const doc = new DOMParser().parseFromString(content)
  style(doc)
  removeTagsPlaceholders(doc)
  const res = new XMLSerializer().serializeToString(doc)
  return res
}
