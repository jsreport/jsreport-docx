const path = require('path')
const { nodeListToArray } = require('../utils')

module.exports = (files) => {
  for (const f of files.filter(f => f.path.endsWith('.xml'))) {
    processEndnotes(files, f)
    processHyperlinks(files, f)
    processFootnotes(files, f)
  }
}

function processEndnotes (files, currentFile) {
  const doc = currentFile.doc
  const endnoteReferenceElements = doc.getElementsByTagName('w:endnoteReference')
  let docEndnotes = files.find(f => f.path === 'word/endnotes.xml')

  if (!docEndnotes) {
    return
  }

  docEndnotes = docEndnotes.doc

  for (let i = 0; i < endnoteReferenceElements.length; i++) {
    const el = endnoteReferenceElements[i]
    const endnoteNode = el.nextSibling
    const clonedEndnoteNode = endnoteNode.cloneNode(true)
    const endnotesNode = docEndnotes.getElementsByTagName('w:endnotes')[0]
    const endnotesNodes = nodeListToArray(endnotesNode.getElementsByTagName('w:endnote'))

    const newId = `${Math.max(...endnotesNodes.map((n) => parseInt(n.getAttribute('w:id'), 10))) + 1}`

    clonedEndnoteNode.setAttribute('w:id', newId)
    el.setAttribute('w:id', newId)

    endnotesNode.appendChild(clonedEndnoteNode)
    el.parentNode.removeChild(endnoteNode)
  }
}

function processHyperlinks (files, currentFile) {
  const doc = currentFile.doc
  const hyperlinkElements = doc.getElementsByTagName('w:hyperlink')
  let docRels = files.find(f => f.path === `word/_rels/${path.basename(currentFile.path)}.rels`)

  if (!docRels) {
    return
  }

  docRels = docRels.doc

  for (let i = 0; i < hyperlinkElements.length; i++) {
    const el = hyperlinkElements[i]
    const relationshipNode = el.firstChild
    const clonedRelationshipNode = relationshipNode.cloneNode()
    const relationshipsNode = docRels.getElementsByTagName('Relationships')[0]
    const relationsNodes = nodeListToArray(relationshipsNode.getElementsByTagName('Relationship'))

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

    clonedRelationshipNode.setAttribute('Id', newId)
    el.setAttribute('r:id', newId)

    relationshipsNode.appendChild(clonedRelationshipNode)
    el.removeChild(relationshipNode)
  }
}

function processFootnotes (files, currentFile) {
  const doc = currentFile.doc
  const footnoteReferenceElements = doc.getElementsByTagName('w:footnoteReference')
  let docFootnotes = files.find(f => f.path === 'word/footnotes.xml')

  if (!docFootnotes) {
    return
  }

  docFootnotes = docFootnotes.doc

  for (let i = 0; i < footnoteReferenceElements.length; i++) {
    const el = footnoteReferenceElements[i]
    const footnoteNode = el.nextSibling

    if (
      footnoteNode == null ||
      footnoteNode.tagName !== 'w:footnote'
    ) {
      continue
    }

    const clonedFootnoteNode = footnoteNode.cloneNode(true)
    const footnotesNode = docFootnotes.getElementsByTagName('w:footnotes')[0]

    if (!footnotesNode) {
      continue
    }

    const footnotesNodes = nodeListToArray(footnotesNode.getElementsByTagName('w:footnote'))

    let newId = footnotesNodes.reduce((lastId, node) => {
      const nodeId = node.getAttribute('w:id')
      const regExp = /^(-?\d+)$/
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

    newId = `${newId}`

    clonedFootnoteNode.setAttribute('w:id', newId)
    el.setAttribute('w:id', newId)

    footnotesNode.appendChild(clonedFootnoteNode)
    el.parentNode.removeChild(footnoteNode)
  }
}
