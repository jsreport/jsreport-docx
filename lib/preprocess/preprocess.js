const { DOMParser, XMLSerializer } = require('xmldom')
const concatTags = require('./concatTags')
module.exports = (content) => {
  const doc = new DOMParser().parseFromString(content)
  concatTags(doc)
  const res = new XMLSerializer().serializeToString(doc)
  return res
}
