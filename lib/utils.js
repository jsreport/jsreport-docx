const { XMLSerializer } = require('xmldom')

module.exports.contentIsXML = (content) => {
  if (!Buffer.isBuffer(content) && typeof content !== 'string') {
    return false
  }

  const str = content.toString()

  return str.startsWith('<?xml') || (/^\s*<[\s\S]*>/).test(str)
}

module.exports.nodeListToArray = (nodes) => {
  const arr = []
  for (let i = 0; i < nodes.length; i++) {
    arr.push(nodes[i])
  }
  return arr
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

module.exports.serializeXml = (doc) => new XMLSerializer().serializeToString(doc).replace(/xmlns(:[a-z0-9]+)?="" ?/g, '')
