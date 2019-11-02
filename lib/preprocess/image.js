const { nodeListToArray } = require('../utils')

module.exports = (files) => {
  const f = files.find(f => f.path === 'word/document.xml')
  const doc = f.doc
  const elements = doc.getElementsByTagName('w:drawing')

  for (let i = 0; i < elements.length; i++) {
    const elDrawing = elements[i]
    const elsLinkClick = nodeListToArray(elDrawing.getElementsByTagName('a:hlinkClick'))

    elsLinkClick.forEach((el) => {
      const attr = el.getAttributeNode('tooltip')

      if (attr) {
        // we mark here that tooltip attr nodes should be unescaped
        // during xml serialization, we do this because tooltip can contain
        // special characters ",',< (ex: {{docxImage src=src width="2cm"}})
        // that are escaped during xml serialization which in the end breaks handlebars
        // parsing, so we indicate that this node should keep its original value
        f.unescapeNodes = f.unescapeNodes || []
        f.unescapeNodes.push(attr)
      }
    })
  }
}
