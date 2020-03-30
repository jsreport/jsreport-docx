const { nodeListToArray } = require('../utils')

module.exports = (files) => {
  const documentFile = files.find(f => f.path === 'word/document.xml').doc
  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const drawningEls = nodeListToArray(documentFile.getElementsByTagName('w:drawing'))

  drawningEls.forEach((drawningEl) => {
    const chartDrawningEl = drawningEl.getElementsByTagName('c:chart')[0]
    const isChart = chartDrawningEl != null

    if (!isChart) {
      return
    }

    const relsElements = nodeListToArray(relsDoc.getElementsByTagName('Relationship'))
    const chartRId = chartDrawningEl.getAttribute('r:id')
    const chartREl = relsElements.find((r) => r.getAttribute('Id') === chartRId)
    const chartFilename = `word/${chartREl.getAttribute('Target')}`
    const chartFile = files.find(f => f.path === chartFilename)
    const chartDoc = chartFile.doc
    const chartTitleEl = chartDoc.getElementsByTagName('c:title')[0]

    if (!chartTitleEl) {
      return
    }

    const chartTitleElClone = chartTitleEl.cloneNode(true)

    drawningEl.insertBefore(chartTitleElClone, drawningEl.firstChild)

    while (chartTitleEl.firstChild) {
      chartTitleEl.removeChild(chartTitleEl.firstChild)
    }
  })
}
