const docxMultiLineRegex = /{{ *docxMultiLine [^{}]{0,500}}}/
const docxMultiLineFindRegex = /{{ *docxMultiLine/

// the problem is that we want to run multiLine helper as an #each, but we want to iterate over all string line-breaks
// this means we need to put the #docxMultiLine up the tree so it encapsulates whole list item

// we find {{multiLine}} literal in the w:t element and move it up the tree so it is before the first w:p
// to keep xml valid we put the helper call inside <docxRemove> node, so in the end it looks something like
// <docxRemove>{{#docxMultiLine aaa}}</docxRemove>
// <w:p><w:r><w:t>{{.}}</w:t></w:r></w:p>
// <docxRemove<{{/docxMultiLine}}</docxRemove>

module.exports = (files) => {
  for (const f of files.filter(f => f.path.endsWith('.xml'))) {
    const doc = f.doc
    const elements = doc.getElementsByTagName('w:t')

    for (let i = 0; i < elements.length; i++) {
      const el = elements[i]

      if (el.textContent.match(docxMultiLineFindRegex)) {
        const helperCall = el.textContent.match(docxMultiLineRegex)[0]
        const wpElement = el.parentNode.parentNode
        const fakeStartElement = doc.createElement('docxRemove')
        fakeStartElement.textContent = helperCall.replace(docxMultiLineFindRegex, '{{#docxMultiLine')

        wpElement.parentNode.insertBefore(fakeStartElement, wpElement)
        el.textContent = '{{.}}'
        const fakeEndElement = doc.createElement('docxRemove')
        fakeEndElement.textContent = '{{/docxMultiLine}}'
        wpElement.parentNode.insertBefore(fakeEndElement, wpElement.nextSibling)
      }
    }
  }
}
