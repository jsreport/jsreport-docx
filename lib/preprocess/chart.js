const { nodeListToArray } = require('../utils')

module.exports = (files) => {
  const documentFile = files.find(f => f.path === 'word/document.xml').doc
  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const drawningEls = nodeListToArray(documentFile.getElementsByTagName('w:drawing'))

  drawningEls.forEach((drawningEl) => {
    const chartDrawningEl = getValidChartEl(drawningEl)

    if (!chartDrawningEl) {
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

function getValidChartEl (drawningEl) {
  let parentEl = drawningEl.parentNode

  const inlineEl = nodeListToArray(drawningEl.childNodes).find((el) => el.nodeName === 'wp:inline')

  if (!inlineEl) {
    return
  }

  const graphicEl = nodeListToArray(inlineEl.childNodes).find((el) => el.nodeName === 'a:graphic')

  if (!graphicEl) {
    return
  }

  const graphicDataEl = nodeListToArray(graphicEl.childNodes).find(el => el.nodeName === 'a:graphicData')

  if (!graphicDataEl) {
    return
  }

  let chartDrawningEl = nodeListToArray(graphicDataEl.childNodes).find(el => el.nodeName === 'c:chart')

  if (!chartDrawningEl) {
    return
  }

  while (parentEl != null) {
    // ignore charts that are part of Fallback tag
    if (parentEl.nodeName === 'mc:Fallback') {
      chartDrawningEl = null
      break
    }

    parentEl = parentEl.parentNode
  }

  return chartDrawningEl
}
