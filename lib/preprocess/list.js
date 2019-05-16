const regexp = /{{#docxList [^{}]{0,500}}}/

module.exports = (doc) => {
  const elements = doc.getElementsByTagName('w:t')

  let openedDocx = false

  function processClosingTag (doc, el) {
    el.textContent = el.textContent.replace('{{/docxList}}', '')

    const wpElement = el.parentNode.parentNode
    const fakeElement = doc.createElement('docxRemove')
    fakeElement.textContent = '{{/docxList}}'

    wpElement.parentNode.insertBefore(fakeElement, wpElement.nextSibling)
  }

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]

    if (el.textContent.includes('{{/docxList}}') && openedDocx) {
      openedDocx = false
      processClosingTag(doc, el)
    }

    if (el.textContent.includes('{{#docxList')) {
      const helperCall = el.textContent.match(regexp)[0]
      const wpElement = el.parentNode.parentNode
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
