const { DOMParser } = require('xmldom')
const { nodeListToArray, serializeXml } = require('../utils')

module.exports = (files) => {
  const f = files.find(f => f.path === 'word/document.xml')

  if (!f.data.toString().includes('$docxPageBreak')) {
    return
  }

  // TODO this regexp is somehow too slow, it needs an optimization
  f.data = f.data.toString().replace(/<w:p.*>.*?\$docxPageBreak.*?(?=<\/w:p>)<\/w:p>/g, (val) => {
    // need to pass a map of existing xml namespaces because we are going to clone the parsed nodes and insert it somewhere else
    const doc = new DOMParser({ xmlns: { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } }).parseFromString(val)
    const wts = doc.getElementsByTagName('w:t')

    let breakFound = false

    const pageBreakP = doc.createElement('w:p')
    const pageBreakWR = doc.createElement('w:r')
    const pageBreakWBR = doc.createElement('w:br')
    pageBreakWBR.setAttribute('w:type', 'page')
    pageBreakP.appendChild(pageBreakWR)
    pageBreakWR.appendChild(pageBreakWBR)

    for (let wt of nodeListToArray(wts)) {
      if (wt.textContent.includes('$docxPageBreak')) {
        breakFound = true
        const parts = wt.textContent.split('$docxPageBreak')
        wt.textContent = parts[0]

        const clonedWR = wt.parentNode.cloneNode(true)
        const clonedWT = clonedWR.getElementsByTagName('w:t')[0]
        clonedWT.textContent = parts[1]
        pageBreakP.appendChild(clonedWR)
        continue
      }

      if (breakFound) {
        pageBreakP.appendChild(wt.parentNode.cloneNode(true))
      }
    }

    return serializeXml(doc) + serializeXml(pageBreakP)
  })
}
