const { DOMParser } = require('xmldom')
const recursiveStringReplaceAsync = require('../recursiveStringReplaceAsync')
const { nodeListToArray, serializeXml } = require('../utils')

module.exports = async (files) => {
  const documentFile = files.find(f => f.path === 'word/document.xml')

  documentFile.data = await recursiveStringReplaceAsync(
    documentFile.data.toString(),
    '<w:p[^>]*__block_helper_container__="true"[^>]*>',
    '</w:p>',
    'g',
    async (val, content, hasNestedMatch) => {
      const doc = new DOMParser().parseFromString(val)
      const paragraphNode = doc.documentElement

      paragraphNode.removeAttribute('__block_helper_container__')

      const blockTextNodes = nodeListToArray(paragraphNode.getElementsByTagName('w:t')).filter((node) => {
        return node.getAttribute('__block_helper__') === 'true'
      })

      for (const textNode of blockTextNodes) {
        const rNode = textNode.parentNode
        let nextNode = rNode.nextSibling
        let nextRNode

        while (nextNode != null) {
          if (nextNode.nodeName === 'w:r') {
            nextRNode = nextNode
            break
          }

          nextNode = nextNode.nextSibling
        }

        if (nextRNode) {
          const childContentNodesLeft = nodeListToArray(nextRNode.childNodes).filter((node) => {
            return !['w:rPr', 'w:br', 'w:cr'].includes(node.nodeName)
          })

          if (childContentNodesLeft.length === 0) {
            // if there are no more content nodes in the w:r then remove it
            nextRNode.parentNode.removeChild(nextRNode)
          }
        }

        rNode.parentNode.removeChild(rNode)
      }

      const childContentNodesLeft = nodeListToArray(paragraphNode.childNodes).filter((node) => {
        return ['w:r', 'w:fldSimple', 'w:hyperlink'].includes(node.nodeName)
      })

      if (childContentNodesLeft.length === 0) {
        // if there are no more content nodes in the paragraph then remove it
        return ''
      }

      return serializeXml(paragraphNode)
    }
  )
}
