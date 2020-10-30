const { nodeListToArray } = require('../utils')

// FIXME: rename docxRaw to docxRun or docxRawRun?
const regexp = /{{docxRaw [^{}]{0,500}}}/

// the problem is that the {{docxRaw}} literal is in a w:t element, which is supposed to only contain literal text,
// but we want the docxRaw helper to provide a raw XML w:r.
// Word 365 is not bothered by that, but other docx editors can be.
// E.g. Word Online displays a broken table and Libreoffice drops the run altogether.

// we find the {{docxRaw}} literal in the w:t element and move it up the tree so it is in its own w:r

module.exports = (files) => {
  const documentFile = files.find(f => f.path === 'word/document.xml').doc
  const generalTextElements = nodeListToArray(documentFile.getElementsByTagName('w:t'))

  for (const textEl of generalTextElements) {
    // there may be more than one docxRaw in a single w:t
    while (textEl.textContent.includes('{{docxRaw')) {
      const helperCall = textEl.textContent.match(regexp)[0]
      const newNode = documentFile.createElement('docxRemove')
      newNode.textContent = helperCall
      textEl.parentNode.parentNode.insertBefore(newNode, textEl.parentNode.nextSibling)
      textEl.textContent = textEl.textContent.replace(regexp, '')
    }
  }
}
