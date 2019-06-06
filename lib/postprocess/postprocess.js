const { DOMParser, XMLSerializer } = require('xmldom')
const removeTagsPlaceholders = require('./removeTagsPlaceholders')
const style = require('./style')
const image = require('./image')

module.exports = (content, relsDocument, newFiles) => {
  const doc = new DOMParser().parseFromString(content)
  style(doc)
  image(doc, relsDocument, newFiles)
  removeTagsPlaceholders(doc)
  const res = new XMLSerializer().serializeToString(doc)
  return res
}
