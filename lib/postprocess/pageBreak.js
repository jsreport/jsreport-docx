
module.exports = (files) => {
  const f = files.find(f => f.path === 'word/document.xml')
  const doc = f.doc

  const pageBreaks = doc.getElementsByTagName('docxPageBreak')

  for (let i = 0; i < pageBreaks.length; i++) {
    const pageBreak = pageBreaks[i]

    const currentParagraphNode = findParentNode(pageBreak, 'w:p')
    const currentParagraphRowNode = pageBreak.parentNode.parentNode

    if (currentParagraphNode && currentParagraphRowNode) {
      const inlineSiblings = getSiblings(pageBreak)
      const rowSiblings = getSiblings(currentParagraphRowNode)

      const newParagraphNode = createParagraphNode(doc, ({ paragraphRowNode }) => {
        const newPageBreakNode = doc.createElement('w:br')
        newPageBreakNode.setAttribute('w:type', 'page')
        paragraphRowNode.appendChild(newPageBreakNode)
      })

      currentParagraphNode.parentNode.insertBefore(newParagraphNode, currentParagraphNode.nextSibling)

      if (rowSiblings.length > 0) {
        const normalizeParagraphNode = createParagraphNode(doc, ({ paragraphNode, paragraphRowNode }) => {
          if (inlineSiblings.length > 0) {
            const newTextNode = doc.createElement('w:t')

            inlineSiblings.forEach((node) => {
              newTextNode.appendChild(node)
            })

            paragraphRowNode.appendChild(newTextNode)
          } else {
            paragraphNode.removeChild(paragraphRowNode)
          }

          rowSiblings.forEach((node) => {
            paragraphNode.appendChild(node)
          })
        })

        currentParagraphNode.parentNode.insertBefore(normalizeParagraphNode, newParagraphNode.nextSibling)
      } else if (inlineSiblings.length > 0) {
        currentParagraphNode.parentNode.insertBefore(createParagraphNode(doc, ({ paragraphRowNode }) => {
          const newTextNode = doc.createElement('w:t')

          inlineSiblings.forEach((node) => {
            newTextNode.appendChild(node)
          })

          paragraphRowNode.appendChild(newTextNode)
        }), newParagraphNode.nextSibling)
      }
    }

    pageBreak.parentNode.removeChild(pageBreak)
  }
}

function findParentNode (node, tagName) {
  let currentNode = node
  let parentNode

  while (currentNode != null) {
    const node = currentNode.parentNode
    const found = node.tagName === tagName

    currentNode = node

    if (found) {
      parentNode = node
      break
    }
  }

  return parentNode
}

function getSiblings (node) {
  const siblings = []
  let currentNode = node

  while (currentNode != null) {
    const node = currentNode.nextSibling
    const found = node != null

    currentNode = node

    if (found) {
      siblings.push(node)
    }
  }

  return siblings
}

function createParagraphNode (doc, onNodeCb) {
  const newParagraphNode = doc.createElement('w:p')
  const newParagraphRowNode = doc.createElement('w:r')

  newParagraphNode.appendChild(newParagraphRowNode)

  onNodeCb && onNodeCb({
    doc,
    paragraphNode: newParagraphNode,
    paragraphRowNode: newParagraphRowNode
  })

  return newParagraphNode
}
