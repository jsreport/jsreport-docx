const { DOMParser, XMLSerializer } = require('xmldom')
const concatTags = require('./concatTags')
const list = require('./list')
const table = require('./table')

module.exports = (content) => {
  const doc = new DOMParser().parseFromString(content)
  concatTags(doc)
  list(doc)
  table(doc)
  let res = new XMLSerializer().serializeToString(doc)
  res = res.replace(/<docxRemove>/g, '').replace(/<\/docxRemove>/g, '')
  return res
}
