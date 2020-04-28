const { nodeListToArray } = require('../utils')

module.exports = function processChart (files, drawingEl) {
  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const chartDrawingEl = getChartEl(drawingEl)

  if (!chartDrawingEl) {
    return
  }

  const relsElements = nodeListToArray(relsDoc.getElementsByTagName('Relationship'))
  const chartRId = chartDrawingEl.getAttribute('r:id')
  const chartREl = relsElements.find((r) => r.getAttribute('Id') === chartRId)
  const chartFilename = `word/${chartREl.getAttribute('Target')}`
  const chartFile = files.find(f => f.path === chartFilename)
  const chartDoc = chartFile.doc
  const chartTitleEl = chartDoc.getElementsByTagName('c:title')[0]

  if (!chartTitleEl) {
    return
  }

  const chartTitleElClone = chartTitleEl.cloneNode(true)

  drawingEl.insertBefore(chartTitleElClone, drawingEl.firstChild)

  while (chartTitleEl.firstChild) {
    chartTitleEl.removeChild(chartTitleEl.firstChild)
  }
}

function getChartEl (drawingEl) {
  let parentEl = drawingEl.parentNode

  const inlineEl = nodeListToArray(drawingEl.childNodes).find((el) => el.nodeName === 'wp:inline')

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

  let chartDrawingEl = nodeListToArray(graphicDataEl.childNodes).find(el => el.nodeName === 'c:chart')

  if (!chartDrawingEl) {
    return
  }

  while (parentEl != null) {
    // ignore charts that are part of Fallback tag
    if (parentEl.nodeName === 'mc:Fallback') {
      chartDrawingEl = null
      break
    }

    parentEl = parentEl.parentNode
  }

  return chartDrawingEl
}
