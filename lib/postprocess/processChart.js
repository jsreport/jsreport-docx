const { DOMParser, XMLSerializer } = require('xmldom')
const { serializeXml, nodeListToArray, getNewRelIdFromBaseId } = require('../utils')

module.exports = async function processChart (files, drawingEl, originalChartsXMLMap, newRelIdCounterMap) {
  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const relsEl = relsDoc.getElementsByTagName('Relationships')[0]
  const contentTypesDoc = files.find(f => f.path === '[Content_Types].xml').doc

  const chartDrawningEl = getValidChartEl(drawingEl)

  if (!chartDrawningEl) {
    return
  }

  let chartRId = chartDrawningEl.getAttribute('r:id')
  let chartREl = nodeListToArray(relsDoc.getElementsByTagName('Relationship')).find((r) => r.getAttribute('Id') === chartRId)
  let chartFilename = `word/${chartREl.getAttribute('Target')}`
  let chartFile = files.find(f => f.path === chartFilename)
  // take the original (not modifed) document
  let chartDoc = originalChartsXMLMap.has(chartFilename) ? new DOMParser().parseFromString(originalChartsXMLMap.get(chartFilename)) : chartFile.doc

  if (!originalChartsXMLMap.has(chartFilename)) {
    originalChartsXMLMap.set(chartFilename, new XMLSerializer().serializeToString(chartDoc))
  }

  let chartRelsFilename = `word/charts/_rels/${chartFilename.split('/').slice(-1)[0]}.rels`
  // take the original (not modifed) document
  let chartRelsDoc = originalChartsXMLMap.has(chartRelsFilename) ? new DOMParser().parseFromString(originalChartsXMLMap.get(chartRelsFilename)) : files.find(f => f.path === chartRelsFilename).doc

  if (!originalChartsXMLMap.has(chartRelsFilename)) {
    originalChartsXMLMap.set(chartRelsFilename, new XMLSerializer().serializeToString(chartRelsDoc))
  }

  const chartStyleRelNode = nodeListToArray(chartRelsDoc.getElementsByTagName('Relationship')).find((el) => {
    return el.getAttribute('Type') === 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
  })

  let chartStyleRelFilename

  if (chartStyleRelNode) {
    chartStyleRelFilename = `word/charts/${chartStyleRelNode.getAttribute('Target')}`
  }

  if (chartStyleRelFilename && !originalChartsXMLMap.has(chartStyleRelFilename)) {
    originalChartsXMLMap.set(chartStyleRelFilename, new XMLSerializer().serializeToString(
      files.find((f) => f.path === chartStyleRelFilename).doc
    ))
  }

  const chartColorStyleRelNode = nodeListToArray(chartRelsDoc.getElementsByTagName('Relationship')).find((el) => {
    return el.getAttribute('Type') === 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
  })

  let chartColorStyleRelFilename

  if (chartColorStyleRelNode) {
    chartColorStyleRelFilename = `word/charts/${chartColorStyleRelNode.getAttribute('Target')}`
  }

  if (chartColorStyleRelFilename && !originalChartsXMLMap.has(chartColorStyleRelFilename)) {
    originalChartsXMLMap.set(chartColorStyleRelFilename, new XMLSerializer().serializeToString(
      files.find((f) => f.path === chartColorStyleRelFilename).doc
    ))
  }

  if (drawingEl.firstChild.nodeName === 'c:title') {
    const newChartTitleEl = drawingEl.firstChild
    const newChartRelId = getNewRelIdFromBaseId(relsDoc, newRelIdCounterMap, chartRId)

    if (chartRId !== newChartRelId) {
      const newRel = nodeListToArray(relsDoc.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Id') === chartRId
      }).cloneNode()

      newRel.setAttribute('Id', newChartRelId)

      const newChartId = files.filter((f) => /word\/charts\/chart(\d+)\.xml/.test(f.path)).reduce((lastId, f) => {
        const numStr = /word\/charts\/chart(\d+)\.xml/.exec(f.path)[1]
        const num = parseInt(numStr, 10)

        if (num > lastId) {
          return num
        }

        return lastId
      }, 0) + 1

      newRel.setAttribute('Target', `charts/chart${newChartId}.xml`)
      relsEl.appendChild(newRel)

      const originalChartXMLStr = originalChartsXMLMap.get(chartFilename)
      const newChartDoc = new DOMParser().parseFromString(originalChartXMLStr)

      chartDoc = newChartDoc

      files.push({
        path: `word/charts/chart${newChartId}.xml`,
        data: originalChartXMLStr,
        // creates new doc
        doc: newChartDoc
      })

      const originalChartRelsXMLStr = originalChartsXMLMap.get(chartRelsFilename)
      const newChartRelsDoc = new DOMParser().parseFromString(originalChartRelsXMLStr)

      files.push({
        path: `word/charts/_rels/chart${newChartId}.xml.rels`,
        data: originalChartRelsXMLStr,
        // creates new doc
        doc: newChartRelsDoc
      })

      let newChartStyleId

      if (chartStyleRelFilename != null) {
        newChartStyleId = files.filter((f) => /word\/charts\/style(\d+)\.xml/.test(f.path)).reduce((lastId, f) => {
          const numStr = /word\/charts\/style(\d+)\.xml/.exec(f.path)[1]
          const num = parseInt(numStr, 10)

          if (num > lastId) {
            return num
          }

          return lastId
        }, 0) + 1

        files.push({
          path: `word/charts/style${newChartStyleId}.xml`,
          data: originalChartsXMLMap.get(chartStyleRelFilename),
          doc: new DOMParser().parseFromString(originalChartsXMLMap.get(chartStyleRelFilename))
        })
      }

      let newChartColorStyleId

      if (chartColorStyleRelFilename != null) {
        newChartColorStyleId = files.filter((f) => /word\/charts\/colors(\d+)\.xml/.test(f.path)).reduce((lastId, f) => {
          const numStr = /word\/charts\/colors(\d+)\.xml/.exec(f.path)[1]
          const num = parseInt(numStr, 10)

          if (num > lastId) {
            return num
          }

          return lastId
        }, 0) + 1

        files.push({
          path: `word/charts/colors${newChartColorStyleId}.xml`,
          data: originalChartsXMLMap.get(chartColorStyleRelFilename),
          doc: new DOMParser().parseFromString(originalChartsXMLMap.get(chartColorStyleRelFilename))
        })
      }

      const newChartType = nodeListToArray(contentTypesDoc.getElementsByTagName('Override')).find((el) => {
        return el.getAttribute('PartName') === `/${chartFilename}`
      }).cloneNode()

      newChartType.setAttribute('PartName', `/word/charts/chart${newChartId}.xml`)

      let newChartStyleType

      if (chartStyleRelFilename != null && newChartStyleId != null) {
        newChartStyleType = nodeListToArray(contentTypesDoc.getElementsByTagName('Override')).find((el) => {
          return el.getAttribute('PartName') === `/${chartStyleRelFilename}`
        }).cloneNode()

        newChartStyleType.setAttribute('PartName', `/word/charts/style${newChartStyleId}.xml`)
      }

      let newChartColorStyleType

      if (chartColorStyleRelFilename && newChartColorStyleId != null) {
        newChartColorStyleType = nodeListToArray(contentTypesDoc.getElementsByTagName('Override')).find((el) => {
          return el.getAttribute('PartName') === `/${chartColorStyleRelFilename}`
        }).cloneNode()

        newChartColorStyleType.setAttribute('PartName', `/word/charts/colors${newChartColorStyleId}.xml`)
      }

      nodeListToArray(newChartRelsDoc.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Type') === 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
      }).setAttribute('Target', `style${newChartStyleId}.xml`)

      nodeListToArray(newChartRelsDoc.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Type') === 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
      }).setAttribute('Target', `colors${newChartColorStyleId}.xml`)

      contentTypesDoc.documentElement.appendChild(newChartType)

      if (newChartStyleType) {
        contentTypesDoc.documentElement.appendChild(newChartStyleType)
      }

      if (newChartColorStyleType) {
        contentTypesDoc.documentElement.appendChild(newChartColorStyleType)
      }
    }

    newChartTitleEl.parentNode.removeChild(newChartTitleEl)

    chartDrawningEl.setAttribute('r:id', newChartRelId)

    const existingChartTitleEl = chartDoc.getElementsByTagName('c:title')[0]

    existingChartTitleEl.parentNode.replaceChild(newChartTitleEl, existingChartTitleEl)
  }

  chartRId = chartDrawningEl.getAttribute('r:id')
  chartREl = nodeListToArray(relsDoc.getElementsByTagName('Relationship')).find((r) => r.getAttribute('Id') === chartRId)
  chartFilename = `word/${chartREl.getAttribute('Target')}`
  chartFile = files.find(f => f.path === chartFilename)
  chartDoc = chartFile.doc
  chartRelsFilename = `word/charts/_rels/${chartFilename.split('/').slice(-1)[0]}.rels`
  chartRelsDoc = files.find(f => f.path === chartRelsFilename).doc

  const chartTitleEl = chartDoc.getElementsByTagName('c:title')[0]

  if (!chartTitleEl) {
    return serializeXml(drawingEl)
  }

  const chartTitleTextElements = nodeListToArray(chartTitleEl.getElementsByTagName('a:t'))

  for (const chartTitleTextEl of chartTitleTextElements) {
    const textContent = chartTitleTextEl.textContent

    if (!textContent.includes('$docxChart')) {
      continue
    }

    const match = textContent.match(/\$docxChart([^$]*)\$/)
    const chartConfig = JSON.parse(Buffer.from(match[1], 'base64').toString())

    // remove chart helper text
    chartTitleTextEl.textContent = chartTitleTextEl.textContent.replace(match[0], '')

    const externalDataEl = chartDoc.getElementsByTagName('c:externalData')[0]

    if (externalDataEl) {
      const externalDataId = externalDataEl.getAttribute('r:id')
      // remove external data reference if exists
      externalDataEl.parentNode.removeChild(externalDataEl)

      const externalXlsxRel = nodeListToArray(chartRelsDoc.getElementsByTagName('Relationship')).find((r) => {
        return r.getAttribute('Id') === externalDataId
      })

      if (externalXlsxRel) {
        const externalXlsxFilename = externalXlsxRel.getAttribute('Target').split('/').slice(-1)[0]
        const externalXlsxFileIndex = files.findIndex((f) => f.path === `word/embeddings/${externalXlsxFilename}`)

        if (externalXlsxFileIndex !== -1) {
          files.splice(externalXlsxFileIndex, 1)
        }

        externalXlsxRel.parentNode.removeChild(externalXlsxRel)
      }
    }

    const chartTypeContentEl = chartDoc.getElementsByTagName('c:plotArea')[0].getElementsByTagName('c:layout')[0].nextSibling
    const supportedCharts = ['areaChart', 'area3DChart', 'barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'pieChart', 'pie3DChart', 'doughnutChart']

    if (!supportedCharts.includes(chartTypeContentEl.localName)) {
      throw new Error(`"${chartTypeContentEl.localName}" type is not supported`)
    }

    const existingChartSeriesElements = nodeListToArray(chartDoc.getElementsByTagName('c:ser'))

    existingChartSeriesElements.forEach((seriesEl) => {
      seriesEl.parentNode.removeChild(seriesEl)
    })

    const placeholderEl = chartDoc.createElement('docxChartSerReplace')

    placeholderEl.textContent = 'sample'

    chartTypeContentEl.insertBefore(placeholderEl, chartTypeContentEl.firstChild)

    chartFile.data = serializeXml(chartFile.doc)
    chartFile.serializeFromDoc = false

    const newChartSeriesElements = chartConfig.data.datasets.map((dataset, datasetIdx) => {
      return `
        <c:ser>
          <c:idx val="${datasetIdx}" />
          <c:order val="${datasetIdx}" />
          <c:tx>
            <c:v>${dataset.label}</c:v>
          </c:tx>
          <c:marker>
            <c:symbol val="none" />
          </c:marker>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="${chartConfig.data.labels.length}" />
                ${chartConfig.data.labels.map((dataLabel, dataLabelIdx) => (`
                  <c:pt idx="${dataLabelIdx}">
                    <c:v>${dataLabel}</c:v>
                  </c:pt>
                `)).join('\n')}
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="${dataset.data.length}" />
                ${dataset.data.map((dataItem, dataItemIdx) => (`
                  <c:pt idx="${dataItemIdx}">
                    <c:v>${dataItem}</c:v>
                  </c:pt>
                `)).join('\n')}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      `
    })

    chartFile.data = chartFile.data.replace(/<docxChartSerReplace[^>]*>[^]*?(?=<\/docxChartSerReplace>)<\/docxChartSerReplace>/g, newChartSeriesElements.join('\n'))
  }

  return serializeXml(drawingEl)
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
