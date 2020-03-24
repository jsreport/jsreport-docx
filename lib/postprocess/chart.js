const stringReplaceAsync = require('string-replace-async')
const { DOMParser } = require('xmldom')
const { serializeXml, nodeListToArray } = require('../utils')

module.exports = async (files) => {
  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const relsElements = nodeListToArray(relsDoc.getElementsByTagName('Relationship'))
  const documentFile = files.find(f => f.path === 'word/document.xml')

  documentFile.data = await stringReplaceAsync(
    documentFile.data.toString(),
    /<a:graphic[^>]*>[^]*?(?=<\/a:graphic>)<\/a:graphic>/g,
    async (val) => {
      const elGraphic = new DOMParser().parseFromString(val)
      const elChart = elGraphic.getElementsByTagName('c:chart')[0]

      if (elChart) {
        const chartRId = elChart.getAttribute('r:id')
        const chartREl = relsElements.find((r) => r.getAttribute('Id') === chartRId)
        const chartFilename = `word/${chartREl.getAttribute('Target')}`
        const chartFile = files.find(f => f.path === chartFilename)

        chartFile.serializeFromDoc = false

        const chartDoc = chartFile.doc
        const chartRelsDoc = files.find(f => f.path === `word/charts/_rels/${chartFilename.split('/').slice(-1)}.rels`).doc
        const chartTitleTextElements = nodeListToArray(chartDoc.getElementsByTagName('c:title')[0].getElementsByTagName('a:t'))

        for (const chartTitleTextEl of chartTitleTextElements) {
          const textContent = chartTitleTextEl.textContent

          if (!textContent.includes('$docxChart')) {
            continue
          }

          const match = textContent.match(/\$docxChart([^$]*)\$/)
          const chartConfig = JSON.parse(Buffer.from(match[1], 'base64').toString())

          // remove chart helper text
          chartTitleTextEl.textContent = ''

          const externalDataEl = chartDoc.getElementsByTagName('c:externalData')[0]

          if (externalDataEl) {
            const externalDataId = externalDataEl.getAttribute('r:id')
            // remove external data reference if exists
            externalDataEl.parentNode.removeChild(externalDataEl)

            const externalXlsxRel = nodeListToArray(chartRelsDoc.getElementsByTagName('Relationship')).find((r) => {
              return r.getAttribute('Id') === externalDataId
            })

            const externalXlsxFilename = externalXlsxRel.getAttribute('Target').split('/').slice(-1)
            const externalXlsxFileIndex = files.findIndex((f) => f.path === `word/embeddings/${externalXlsxFilename}`)

            files.splice(externalXlsxFileIndex, 1)

            externalXlsxRel.parentNode.removeChild(externalXlsxRel)
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
      }

      return serializeXml(elGraphic)
    }
  )
}
