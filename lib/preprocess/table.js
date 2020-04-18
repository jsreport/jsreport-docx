
const regexp = /{{#docxTable [^{}]{0,500}}}/

function processClosingTag (doc, el) {
  el.textContent = el.textContent.replace('{{/docxTable}}', '')

  const wpElement = el.parentNode.parentNode.parentNode.parentNode
  const fakeElement = doc.createElement('docxRemove')
  fakeElement.textContent = '{{/docxTable}}'

  wpElement.parentNode.insertBefore(fakeElement, wpElement.nextSibling)
}

// the same idea as list, check the docs there
module.exports = (files) => {
  for (const f of files.filter(f => f.path.endsWith('.xml'))) {
    const doc = f.doc
    const elements = doc.getElementsByTagName('w:t')
    let openTags = 0

    for (let i = 0; i < elements.length; i++) {
      const el = elements[i]

      if (el.textContent.includes('{{/docxTable}}') && openTags > 0) {
        openTags--
        processClosingTag(doc, el)
      }

      if (el.textContent.includes('{{#docxTable')) {
        openTags++
        const helperCall = el.textContent.match(regexp)[0]
        const wpElement = el.parentNode.parentNode.parentNode.parentNode
        const fakeElement = doc.createElement('docxRemove')
        fakeElement.textContent = helperCall

        wpElement.parentNode.insertBefore(fakeElement, wpElement)
        el.textContent = el.textContent.replace(regexp, '')

        if (el.textContent.includes('{{/docxTable')) {
          openTags--
          processClosingTag(doc, el)
        }
      }
    }
  }
}
