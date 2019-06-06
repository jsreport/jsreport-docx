
const regexp = /{{#docxTable [^{}]{0,500}}}/

// the same idea as list, check the docs there
module.exports = (files) => {
  const doc = files.find(f => f.path === 'word/document.xml').doc
  const elements = doc.getElementsByTagName('w:t')

  let openedDocx = false

  function processClosingTag (doc, el) {
    el.textContent = el.textContent.replace('{{/docxTable}}', '')

    const wpElement = el.parentNode.parentNode.parentNode.parentNode
    const fakeElement = doc.createElement('docxRemove')
    fakeElement.textContent = '{{/docxTable}}'

    wpElement.parentNode.insertBefore(fakeElement, wpElement.nextSibling)
  }

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]

    if (el.textContent.includes('{{/docxTable}}') && openedDocx) {
      openedDocx = false
      processClosingTag(doc, el)
    }

    if (el.textContent.includes('{{#docxTable')) {
      const helperCall = el.textContent.match(regexp)[0]
      const wpElement = el.parentNode.parentNode.parentNode.parentNode
      const fakeElement = doc.createElement('docxRemove')
      fakeElement.textContent = helperCall

      wpElement.parentNode.insertBefore(fakeElement, wpElement)
      el.textContent = el.textContent.replace(regexp, '')
      if (el.textContent.includes('{{/docxList')) {
        processClosingTag(doc, el)
      } else {
        openedDocx = true
      }
    }
  }
}
