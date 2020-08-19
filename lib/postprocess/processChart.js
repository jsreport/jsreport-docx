const { DOMParser, XMLSerializer } = require('xmldom')
const moment = require('moment')
const toExcelDate = require('js-excel-date-convert').toExcelDate
const { serializeXml, nodeListToArray, getChartEl, getNewRelIdFromBaseId } = require('../utils')

module.exports = async function processChart (files, drawingEl, originalChartsXMLMap, newRelIdCounterMap) {
  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const relsEl = relsDoc.getElementsByTagName('Relationships')[0]
  const contentTypesDoc = files.find(f => f.path === '[Content_Types].xml').doc

  const chartDrawingEl = getChartEl(drawingEl)

  if (!chartDrawingEl) {
    return
  }

  let chartRId = chartDrawingEl.getAttribute('r:id')
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

  if (drawingEl.firstChild.nodeName === `${chartDrawingEl.prefix}:title`) {
    const newChartTitleEl = drawingEl.firstChild
    const newChartRelId = getNewRelIdFromBaseId(relsDoc, newRelIdCounterMap, chartRId)

    if (chartRId !== newChartRelId) {
      const newRel = nodeListToArray(relsDoc.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Id') === chartRId
      }).cloneNode()

      newRel.setAttribute('Id', newChartRelId)

      let getIdRegexp

      if (chartDrawingEl.prefix === 'cx') {
        getIdRegexp = () => /word\/charts\/chartEx(\d+)\.xml/
      } else {
        getIdRegexp = () => /word\/charts\/chart(\d+)\.xml/
      }

      const newChartId = files.filter((f) => getIdRegexp().test(f.path)).reduce((lastId, f) => {
        const numStr = getIdRegexp().exec(f.path)[1]
        const num = parseInt(numStr, 10)

        if (num > lastId) {
          return num
        }

        return lastId
      }, 0) + 1

      let filePrefix = 'chart'

      if (chartDrawingEl.prefix === 'cx') {
        filePrefix = 'chartEx'
      }

      newRel.setAttribute('Target', `charts/${filePrefix}${newChartId}.xml`)

      relsEl.appendChild(newRel)

      const originalChartXMLStr = originalChartsXMLMap.get(chartFilename)
      const newChartDoc = new DOMParser().parseFromString(originalChartXMLStr)

      chartDoc = newChartDoc

      files.push({
        path: `word/charts/${filePrefix}${newChartId}.xml`,
        data: originalChartXMLStr,
        // creates new doc
        doc: newChartDoc
      })

      const originalChartRelsXMLStr = originalChartsXMLMap.get(chartRelsFilename)
      const newChartRelsDoc = new DOMParser().parseFromString(originalChartRelsXMLStr)

      files.push({
        path: `word/charts/_rels/${filePrefix}${newChartId}.xml.rels`,
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

      newChartType.setAttribute('PartName', `/word/charts/${filePrefix}${newChartId}.xml`)

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

    chartDrawingEl.setAttribute('r:id', newChartRelId)

    const existingChartTitleEl = chartDoc.getElementsByTagName(`${chartDrawingEl.prefix}:title`)[0]

    existingChartTitleEl.parentNode.replaceChild(newChartTitleEl, existingChartTitleEl)
  }

  chartRId = chartDrawingEl.getAttribute('r:id')
  chartREl = nodeListToArray(relsDoc.getElementsByTagName('Relationship')).find((r) => r.getAttribute('Id') === chartRId)
  chartFilename = `word/${chartREl.getAttribute('Target')}`
  chartFile = files.find(f => f.path === chartFilename)
  chartDoc = chartFile.doc
  chartRelsFilename = `word/charts/_rels/${chartFilename.split('/').slice(-1)[0]}.rels`
  chartRelsDoc = files.find(f => f.path === chartRelsFilename).doc

  const chartTitleEl = chartDoc.getElementsByTagName(`${chartDrawingEl.prefix}:title`)[0]

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

    const externalDataEl = chartDoc.getElementsByTagName(`${chartDrawingEl.prefix}:externalData`)[0]

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

    if (chartDrawingEl.prefix === 'cx') {
      const chartSeriesEl = chartDoc.getElementsByTagName('cx:plotArea')[0].getElementsByTagName('cx:series')[0]
      const chartType = chartSeriesEl.getAttribute('layoutId')
      const supportedCharts = ['waterfall', 'treemap', 'sunburst', 'funnel', 'clusteredColumn']

      if (!supportedCharts.includes(chartType)) {
        throw new Error(`"${chartType}" type (chartEx) is not supported`)
      }

      const chartDataEl = chartDoc.getElementsByTagName('cx:chartData')[0]
      const existingDataItemsElements = nodeListToArray(chartDataEl.getElementsByTagName('cx:data'))
      const dataPlaceholderEl = chartDoc.createElement('docxChartexDataReplace')
      const seriesPlaceholderEl = chartDoc.createElement('docxChartexSeriesReplace')

      dataPlaceholderEl.textContent = 'sample'
      seriesPlaceholderEl.textContent = 'sample'

      chartDataEl.appendChild(dataPlaceholderEl)
      chartSeriesEl.parentNode.insertBefore(seriesPlaceholderEl, chartSeriesEl.nextSibling)

      existingDataItemsElements.forEach((dataItemEl) => {
        dataItemEl.parentNode.removeChild(dataItemEl)
      })

      chartSeriesEl.parentNode.removeChild(chartSeriesEl)

      chartFile.data = serializeXml(chartFile.doc)
      chartFile.serializeFromDoc = false

      let newDataItemElement = existingDataItemsElements[0].cloneNode(true)

      newDataItemElement.setAttribute('id', 0)

      addChartexItem(chartDoc, {
        name: 'cx:strDim',
        type: chartType,
        data: Array.isArray(chartConfig.data.labels[0]) ? chartConfig.data.labels.map((subLabels) => ({ items: subLabels })) : [{ items: chartConfig.data.labels }]
      }, newDataItemElement)

      addChartexItem(chartDoc, { name: 'cx:numDim', type: chartType, data: [{ items: chartConfig.data.datasets[0].data || [] }] }, newDataItemElement)

      let newChartSeriesElement = chartSeriesEl.cloneNode(true)

      addChartexItem(chartDoc, { name: 'cx:tx', data: chartConfig.data.datasets[0].label || '' }, newChartSeriesElement)
      addChartexItem(chartDoc, { name: 'cx:dataId', data: newDataItemElement.getAttribute('id') }, newChartSeriesElement)

      newDataItemElement = serializeXml(newDataItemElement)
      newChartSeriesElement = serializeXml(newChartSeriesElement)

      chartFile.data = chartFile.data.replace(/<docxChartexDataReplace[^>]*>[^]*?(?=<\/docxChartexDataReplace>)<\/docxChartexDataReplace>/g, newDataItemElement)
      chartFile.data = chartFile.data.replace(/<docxChartexSeriesReplace[^>]*>[^]*?(?=<\/docxChartexSeriesReplace>)<\/docxChartexSeriesReplace>/g, newChartSeriesElement)
    } else {
      const chartTypeContentEl = chartDoc.getElementsByTagName('c:plotArea')[0].getElementsByTagName('c:layout')[0].nextSibling
      const chartType = chartTypeContentEl.localName
      const supportedCharts = [
        'areaChart', 'area3DChart', 'barChart', 'bar3DChart', 'lineChart', 'line3DChart',
        'pieChart', 'pie3DChart', 'doughnutChart', 'stockChart', 'scatterChart', 'bubbleChart'
      ]

      if (!supportedCharts.includes(chartType)) {
        throw new Error(`"${chartType}" type is not supported`)
      }

      const existingChartSeriesElements = nodeListToArray(chartDoc.getElementsByTagName('c:ser'))

      const placeholderEl = chartDoc.createElement('docxChartSerReplace')

      placeholderEl.textContent = 'sample'

      let placeholderRefNode = chartTypeContentEl.firstChild

      if (existingChartSeriesElements.length > 0) {
        placeholderRefNode = existingChartSeriesElements[0]
      } else {
        for (const childNode of nodeListToArray(chartTypeContentEl.childNodes)) {
          if (childNode.nodeName === 'c:dLbls') {
            placeholderRefNode = childNode
            break
          }
        }
      }

      chartTypeContentEl.insertBefore(placeholderEl, placeholderRefNode)

      existingChartSeriesElements.forEach((seriesEl) => {
        seriesEl.parentNode.removeChild(seriesEl)
      })

      chartFile.data = serializeXml(chartFile.doc)
      chartFile.serializeFromDoc = false

      const newChartSeriesElements = chartConfig.data.datasets.map((dataset, datasetIdx) => {
        const newChartSerieNode = existingChartSeriesElements[datasetIdx].cloneNode(true)

        removeChildNodes('c:extLst', newChartSerieNode)

        addChartSerieItem(chartDoc, { name: 'c:idx', data: datasetIdx }, newChartSerieNode)
        addChartSerieItem(chartDoc, { name: 'c:order', data: datasetIdx }, newChartSerieNode)
        addChartSerieItem(chartDoc, { name: 'c:tx', data: [dataset.label] }, newChartSerieNode)

        if (chartType === 'scatterChart' || chartType === 'bubbleChart') {
          addChartSerieItem(chartDoc, { name: 'c:xVal', data: chartConfig.data.labels }, newChartSerieNode)

          if (chartType === 'bubbleChart') {
            if (dataset.data.some((d) => !Array.isArray(d))) {
              throw new Error('bubbleChart expects each data item to be array of [yValue, sizeValue]')
            }

            addChartSerieItem(chartDoc, { name: 'c:yVal', data: dataset.data.map((d) => d[0]) }, newChartSerieNode)
            addChartSerieItem(chartDoc, { name: 'c:bubbleSize', data: dataset.data.map((d) => d[1]) }, newChartSerieNode)
          } else {
            addChartSerieItem(chartDoc, { name: 'c:yVal', data: dataset.data }, newChartSerieNode)
          }
        } else {
          addChartSerieItem(chartDoc, { name: 'c:cat', type: chartType, data: chartConfig.data.labels }, newChartSerieNode)
          addChartSerieItem(chartDoc, { name: 'c:val', data: dataset.data }, newChartSerieNode)
        }

        return serializeXml(newChartSerieNode)
      })

      chartFile.data = chartFile.data.replace(/<docxChartSerReplace[^>]*>[^]*?(?=<\/docxChartSerReplace>)<\/docxChartSerReplace>/g, newChartSeriesElements.join('\n'))
    }
  }

  return serializeXml(drawingEl)
}

function addChartexItem (docNode, nodeInfo, targetNode) {
  let newNode

  const existingNode = findChildNode(nodeInfo.name, targetNode)

  if (existingNode) {
    newNode = existingNode.cloneNode(true)
  } else {
    newNode = docNode.createElement(nodeInfo.name)
  }

  switch (nodeInfo.name) {
    case 'cx:strDim':
    case 'cx:numDim':
      let empty = false
      const isHierarchyType = nodeInfo.type === 'treemap' || nodeInfo.type === 'sunburst'
      const isNum = nodeInfo.name === 'cx:numDim'
      let type = isNum ? 'val' : 'cat'

      if (isNum && isHierarchyType) {
        type = 'size'
      }

      if (!isNum && nodeInfo.type === 'clusteredColumn') {
        empty = true
      }

      newNode.setAttribute('type', type)

      removeChildNodes('cx:f', newNode)

      const existingLvlNodes = findChildNode('cx:lvl', newNode, true)

      if (!empty) {
        let targetData = nodeInfo.data

        if (!isNum && isHierarchyType) {
          targetData = targetData.reverse()
        }

        for (const [idx, lvlInfo] of targetData.entries()) {
          let lvlNode

          if (existingLvlNodes[idx] != null) {
            lvlNode = existingLvlNodes[idx].cloneNode(true)
            newNode.insertBefore(lvlNode, existingLvlNodes[0])
          } else {
            lvlNode = docNode.createElement('cx:lvl')

            if (existingLvlNodes.length > 0) {
              newNode.insertBefore(lvlNode, existingLvlNodes[0])
            } else {
              newNode.appendChild(lvlNode)
            }
          }

          lvlNode.setAttribute('ptCount', lvlInfo.items.length)

          const existingPtNodes = findChildNode('cx:pt', lvlNode, true)

          for (const [itemIdx, item] of lvlInfo.items.entries()) {
            let ptNode

            if (existingPtNodes[itemIdx] != null) {
              ptNode = existingPtNodes[itemIdx].cloneNode(true)
              lvlNode.insertBefore(ptNode, existingPtNodes[0])
            } else {
              ptNode = docNode.createElement('cx:pt')

              if (existingPtNodes.length > 0) {
                lvlNode.insertBefore(ptNode, existingPtNodes[0])
              } else {
                lvlNode.appendChild(ptNode)
              }
            }

            ptNode.setAttribute('idx', itemIdx)
            ptNode.textContent = item != null ? item : ''
          }

          for (const ePtNode of existingPtNodes) {
            ePtNode.parentNode.removeChild(ePtNode)
          }
        }
      } else {
        newNode = null
      }

      for (const eLvlNode of existingLvlNodes) {
        eLvlNode.parentNode.removeChild(eLvlNode)
      }

      break
    case 'cx:tx':
      const txDataNode = findOrCreateChildNode(docNode, 'cx:txData', newNode)

      removeChildNodes('cx:f', txDataNode)

      const txValueNode = findOrCreateChildNode(docNode, 'cx:v', txDataNode)

      txValueNode.textContent = nodeInfo.data

      break
    case 'cx:dataId':
      newNode.setAttribute('val', nodeInfo.data)
      break
    default:
      throw new Error(`node chartex item "${nodeInfo.name}" not supported`)
  }

  if (!newNode) {
    if (existingNode) {
      targetNode.removeChild(existingNode)
    }

    return
  }

  if (existingNode) {
    targetNode.replaceChild(newNode, existingNode)
  } else {
    targetNode.appendChild(newNode)
  }
}

function addChartSerieItem (docNode, nodeInfo, targetNode) {
  let newNode

  const existingNode = findChildNode(nodeInfo.name, targetNode)

  if (existingNode) {
    newNode = existingNode.cloneNode(true)
  } else {
    newNode = docNode.createElement(nodeInfo.name)
  }

  switch (nodeInfo.name) {
    case 'c:idx':
    case 'c:order':
      newNode.setAttribute('val', nodeInfo.data)
      break
    case 'c:tx':
    case 'c:cat':
    case 'c:val':
    case 'c:xVal':
    case 'c:yVal':
    case 'c:bubbleSize':
      const shouldBeDateType = nodeInfo.name === 'c:cat' && nodeInfo.type === 'stockChart'
      let isNum = nodeInfo.name === 'c:val' || nodeInfo.name === 'c:xVal' || nodeInfo.name === 'c:yVal' || nodeInfo.name === 'c:bubbleSize'

      if (shouldBeDateType) {
        isNum = true
      }

      const refNode = findOrCreateChildNode(docNode, isNum ? 'c:numRef' : 'c:strRef', newNode)
      removeChildNodes('c:f', refNode)
      const cacheNode = findOrCreateChildNode(docNode, isNum ? 'c:numCache' : 'c:strCache', refNode)
      const existingFormatNode = findChildNode('c:formatCode', cacheNode)

      if (isNum && !existingFormatNode) {
        const formatNode = docNode.createElement('c:formatCode')
        formatNode.textContent = shouldBeDateType ? 'm/d/yy' : 'General'
        cacheNode.insertBefore(formatNode, cacheNode.firstChild)
      }

      const ptCountNode = findOrCreateChildNode(docNode, 'c:ptCount', cacheNode)

      ptCountNode.setAttribute('val', nodeInfo.data.length)

      const existingPtNodes = findChildNode('c:pt', cacheNode, true)

      for (const [idx, item] of nodeInfo.data.entries()) {
        let ptNode

        if (existingPtNodes[idx] != null) {
          ptNode = existingPtNodes[idx].cloneNode(true)
          cacheNode.insertBefore(ptNode, existingPtNodes[0])
        } else {
          ptNode = docNode.createElement('c:pt')

          if (existingPtNodes.length > 0) {
            cacheNode.insertBefore(ptNode, existingPtNodes[0])
          } else {
            cacheNode.appendChild(ptNode)
          }
        }

        ptNode.setAttribute('idx', idx)

        const ptValueNode = findOrCreateChildNode(docNode, 'c:v', ptNode)

        let value = item

        if (shouldBeDateType) {
          const parsedValue = moment(item)

          if (parsedValue.isValid() === false) {
            throw new Error(`label for "${nodeInfo.type}" should be date string in format of YYYY-MM-DD`)
          }

          value = toExcelDate(parsedValue.toDate())
        }

        ptValueNode.textContent = value
      }

      for (const eNode of existingPtNodes) {
        eNode.parentNode.removeChild(eNode)
      }

      break
    default:
      throw new Error(`node chart item "${nodeInfo.name}" not supported`)
  }

  if (existingNode) {
    targetNode.replaceChild(newNode, existingNode)
  } else {
    targetNode.appendChild(newNode)
  }
}

function findOrCreateChildNode (docNode, nodeName, targetNode) {
  let result
  const existingNode = findChildNode(nodeName, targetNode)

  if (!existingNode) {
    result = docNode.createElement(nodeName)
    targetNode.appendChild(result)
  } else {
    result = existingNode
  }

  return result
}

function findChildNode (nodeName, targetNode, allNodes = false) {
  let result = []

  for (let i = 0; i < targetNode.childNodes.length; i++) {
    let found = false
    const childNode = targetNode.childNodes[i]

    if (childNode.nodeName === nodeName) {
      found = true
      result.push(childNode)
    }

    if (found && !allNodes) {
      break
    }
  }

  return allNodes ? result : result[0]
}

function removeChildNodes (nodeName, targetNode) {
  for (let i = 0; i < targetNode.childNodes.length; i++) {
    const childNode = targetNode.childNodes[i]

    if (childNode.nodeName === nodeName) {
      targetNode.removeChild(childNode)
    }
  }
}
