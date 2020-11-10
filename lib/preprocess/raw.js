const { nodeListToArray } = require('../utils')

const regexp = /{{docxRaw [^{}]{0,500}}}/

const supportedParentElements = ['w:r', 'w:p', 'w:tc']

// the problem is that the {{docxRaw}} literal is in a w:t element, which is supposed to only contain literal text,
// but we want the docxRaw helper to provide a raw XML fragment.
// Word 365 is not bothered by that, but other docx editors can be.
// E.g. Word Online displays a broken table and Libreoffice drops the run altogether.

// we find the {{docxRaw}} literal in the w:t element and move it up the tree so it is in its desired location.

module.exports = (files) => {
  const documentFile = files.find(f => f.path === 'word/document.xml').doc
  const generalTextElements = nodeListToArray(documentFile.getElementsByTagName('w:t'))

  for (const textEl of generalTextElements) {
    // there may be more than one docxRaw helper call in a single w:t
    while (textEl.textContent.includes('{{docxRaw')) {
      const xmlArg = textEl.textContent.match(/xml=(?<xml>[^{}\s]+)/)
      const replaceParentElementArg = textEl.textContent.match(/replaceParentElement=['"](?<replaceParentElement>[^{}\s]+)['"]/)
      if (!xmlArg || !xmlArg.groups || !replaceParentElementArg || !replaceParentElementArg.groups) {
        throw new Error('Expected "xml" and "replaceParentElement" parameters for the docxRaw helper')
      }
      if (!supportedParentElements.includes(replaceParentElementArg.groups.replaceParentElement)) {
        throw new Error('Expected a "replaceParentElement" parameter to be one of ' + supportedParentElements + ', got ' + replaceParentElementArg.groups.replaceParentElement)
      }

      const helperCall = textEl.textContent.match(regexp)[0]
      const newNode = documentFile.createElement('docxRemove')
      newNode.textContent = helperCall

      const refElement = getReferenceElement(textEl, replaceParentElementArg.groups.replaceParentElement)
      // ensure reference element has the proper type, especially useful for w:tc elements which should not be used outside tables cells.
      if (refElement.nodeName !== replaceParentElementArg.groups.replaceParentElement) {
        throw new Error('Reference element does not match replaceParentElement parameter, expected ' + replaceParentElementArg.groups.replaceParentElement + ', got ' + refElement.nodeName)
      }

      // insert the new node right after the reference element
      refElement.parentNode.insertBefore(newNode, refElement.nextSibling)

      // remove the helper from its original location
      textEl.textContent = textEl.textContent.replace(regexp, '')
    }
  }
}

function getReferenceElement (textEl, replaceParentElement) {
  switch (replaceParentElement) {
    case 'w:p': return textEl.parentNode.parentNode
    case 'w:tc': return textEl.parentNode.parentNode.parentNode
    default: return textEl.parentNode
  }
}
