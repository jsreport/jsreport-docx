const { XMLSerializer } = require('xmldom')

function nodeListToArray (nodes) {
  const arr = []
  for (let i = 0; i < nodes.length; i++) {
    arr.push(nodes[i])
  }
  return arr
}

function getNewRelId (relsDoc) {
  const relationsNodes = nodeListToArray(relsDoc.getElementsByTagName('Relationship'))

  let newId = relationsNodes.reduce((lastId, node) => {
    const nodeId = node.getAttribute('Id')
    const regExp = /^rId(\d+)$/
    const match = regExp.exec(nodeId)

    if (!match || !match[1]) {
      return lastId
    }

    const num = parseInt(match[1], 10)

    if (num > lastId) {
      return num
    }

    return lastId
  }, 0) + 1

  newId = `rId${newId}`

  return newId
}

module.exports.contentIsXML = (content) => {
  if (!Buffer.isBuffer(content) && typeof content !== 'string') {
    return false
  }

  const str = content.toString()

  return str.startsWith('<?xml') || (/^\s*<[\s\S]*>/).test(str)
}

module.exports.pxToEMU = (val) => {
  return Math.round(val * 914400 / 96)
}

module.exports.cmToEMU = (val) => {
  // cm to dxa
  const dxa = val * 567.058823529411765
  // dxa to EMU
  return Math.round(dxa * 914400 / 72 / 20)
}

module.exports.getNewRelIdFromBaseId = (relsDoc, itemsMap, baseId) => {
  const counter = itemsMap.get(baseId) || 0

  itemsMap.set(baseId, counter + 1)

  if (counter === 0) {
    return baseId
  }

  return getNewRelId(relsDoc)
}

module.exports.serializeXml = (doc) => new XMLSerializer().serializeToString(doc).replace(/ xmlns(:[a-z0-9]+)?=""/g, '')
module.exports.getNewRelId = getNewRelId
module.exports.nodeListToArray = nodeListToArray
