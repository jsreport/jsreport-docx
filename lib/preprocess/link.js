const { nodeListToArray } = require('../utils')

// hyperlink value is url encoded so developer cannot use handlebars there
// we url decode this cases before handlebars runs
// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%7b%7burl%7d%7d" TargetMode="External"/>
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
    const refId = el.getAttribute('w:id')
    const endnotes = nodeListToArray(docEndnotes.getElementsByTagName('w:endnote'))
    const endnoteEl = endnotes.filter(e => e.getAttribute('w:id') === refId)[0]

    if (!endnoteEl) {
      continue
    }

    el.removeAttribute('w:id')

    const clonedEndnoteEl = endnoteEl.cloneNode(true)

    clonedEndnoteEl.removeAttribute('w:id')

    const fakeElement = doc.createElement('docxRemove')

    fakeElement.appendChild(clonedEndnoteEl)

    el.parentNode.insertBefore(fakeElement, el.nextSibling)

    endnoteEl.parentNode.removeChild(endnoteEl)
  }
}

function processHyperlinks (files, currentFile) {
  const doc = currentFile.doc
  const hyperlinkElements = doc.getElementsByTagName('w:hyperlink')
  let docRels = files.find(f => f.path === 'word/_rels/document.xml.rels')

  if (!docRels) {
    return
  }

  docRels = docRels.doc

  for (let i = 0; i < hyperlinkElements.length; i++) {
    const el = hyperlinkElements[i]
    const relationshipId = el.getAttribute('r:id')
    const rels = nodeListToArray(docRels.getElementsByTagName('Relationships')[0].getElementsByTagName('Relationship'))
    const relationshipEl = rels.filter(r => r.getAttribute('Id') === relationshipId && r.getAttribute('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink')[0]

    if (!relationshipEl) {
      continue
    }

    el.removeAttribute('r:id')

    const clonedRelationshipEl = relationshipEl.cloneNode()

    const decodedTarget = decodeURIComponent(clonedRelationshipEl.getAttribute('Target'))

    if (decodedTarget.includes('{{')) {
      clonedRelationshipEl.setAttribute('Target', decodedTarget)
    }

    clonedRelationshipEl.removeAttribute('Id')

    const fakeElement = doc.createElement('docxRemove')

    fakeElement.appendChild(clonedRelationshipEl)

    el.insertBefore(fakeElement, el.firstChild)

    relationshipEl.parentNode.removeChild(relationshipEl)
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
    const refId = el.getAttribute('w:id')
    const footnotes = nodeListToArray(docFootnotes.getElementsByTagName('w:footnote'))
    const footnoteEl = footnotes.filter(e => e.getAttribute('w:id') === refId)[0]

    if (!footnoteEl) {
      continue
    }

    el.removeAttribute('w:id')

    const clonedFootnoteEl = footnoteEl.cloneNode(true)

    clonedFootnoteEl.removeAttribute('w:id')

    const fakeElement = doc.createElement('docxRemove')

    fakeElement.appendChild(clonedFootnoteEl)

    el.parentNode.insertBefore(fakeElement, el.nextSibling)

    footnoteEl.parentNode.removeChild(footnoteEl)
  }
}
