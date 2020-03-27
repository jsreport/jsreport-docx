const { nodeListToArray } = require('../utils')

module.exports = (files) => {
  const documentFile = files.find(f => f.path === 'word/document.xml').doc
  const drawningEls = nodeListToArray(documentFile.getElementsByTagName('w:drawing'))

  drawningEls.forEach((drawningEl) => {
    const isImg = drawningEl.getElementsByTagName('pic:pic').length > 0

    if (!isImg) {
      return
    }

    const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc

    const elLinkClicks = drawningEl.getElementsByTagName('a:hlinkClick')
    const elLinkClick = elLinkClicks[0]

    if (!elLinkClick) {
      return
    }

    // to support hyperlink generation in a loop, we need to insert the hyperlink definition into
    // the document xml, so when template engine runs it evaluates with the correct data context
    const hyperlinkRelId = elLinkClick.getAttribute('r:id')

    const hyperlinkRelEl = nodeListToArray(relsDoc.getElementsByTagName('Relationship')).find((el) => {
      return el.getAttribute('Id') === hyperlinkRelId
    })

    const hyperlinkRelElClone = hyperlinkRelEl.cloneNode()

    const decodedTarget = decodeURIComponent(hyperlinkRelElClone.getAttribute('Target'))

    if (decodedTarget.includes('{{')) {
      hyperlinkRelElClone.setAttribute('Target', decodedTarget)
    }

    drawningEl.insertBefore(hyperlinkRelElClone, drawningEl.firstChild)
  })
}
