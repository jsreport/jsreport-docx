const { nodeListToArray } = require('../utils')
const regexp = /{{#?docxTable [^{}]{0,500}}}/

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

      if (
        el.textContent.includes('{{docxTable') &&
        el.textContent.includes('rows=') &&
        el.textContent.includes('columns=')
      ) {
        // full table mode
        let helperCall = el.textContent.match(regexp)[0]
        const originalTextContent = el.textContent

        // setting the cell text to be the value for the rows (before we clone)
        el.textContent = el.textContent.replace(regexp, '{{lookup ../this @index}}')

        const cellNode = el.parentNode.parentNode.parentNode
        const rowNode = cellNode.parentNode
        const newRowNode = rowNode.cloneNode(true)

        helperCall = helperCall.replace('{{docxTable', '{{#docxTable')

        // first row, handling the cells for the column names
        el.textContent = originalTextContent.replace(regexp, '{{this}}')

        processOpeningTag(doc, cellNode, helperCall.replace('rows=', 'ignore='))
        processClosingTag(doc, cellNode)

        // row template, handling the cells for the data values
        rowNode.parentNode.insertBefore(newRowNode, rowNode.nextSibling)
        const cellInNewRowNode = nodeListToArray(newRowNode.childNodes).find((node) => node.nodeName === 'w:tc')

        processOpeningTag(doc, cellInNewRowNode, helperCall.replace('rows=', 'ignore=').replace('columns=', 'ignore='))
        processClosingTag(doc, cellInNewRowNode)

        processOpeningTag(doc, newRowNode, helperCall)
        processClosingTag(doc, newRowNode)
      } else if (el.textContent.includes('{{#docxTable')) {
        const helperCall = el.textContent.match(regexp)[0]
        const isVertical = el.textContent.includes('vertical=')
        const isNormal = !isVertical

        if (isNormal) {
          openTags++
        }

        if (isVertical) {
          const cellNode = el.parentNode.parentNode.parentNode
          const cellIndex = getCellIndex(cellNode)
          const [affectedRows, textNodeTableClose] = getNextRowsUntilTableClose(cellNode.parentNode)

          if (textNodeTableClose) {
            textNodeTableClose.textContent = textNodeTableClose.textContent.replace('{{/docxTable}}', '')
          }

          processOpeningTag(doc, el, helperCall, isVertical)
          processClosingTag(doc, el, isVertical)

          for (const rowNode of affectedRows) {
            const cellNodes = nodeListToArray(rowNode.childNodes).filter((node) => node.nodeName === 'w:tc')
            const cellNode = cellNodes[cellIndex]

            if (cellNode) {
              processOpeningTag(doc, cellNode, helperCall, isVertical)
              processClosingTag(doc, cellNode, isVertical)
            }
          }
        } else {
          processOpeningTag(doc, el, helperCall, isVertical)
        }

        if (isNormal && el.textContent.includes('{{/docxTable')) {
          openTags--
          processClosingTag(doc, el)
        }
      }
    }
  }
}

function processOpeningTag (doc, el, helperCall, isVertical = false) {
  if (el.nodeName === 'w:t') {
    el.textContent = el.textContent.replace(regexp, '')
  }

  const fakeElement = doc.createElement('docxRemove')

  fakeElement.textContent = helperCall

  let refElement

  if (el.nodeName !== 'w:t') {
    refElement = el
  } else {
    if (isVertical) {
      // ref is the column w:tc
      refElement = el.parentNode.parentNode.parentNode
    } else {
      // ref is the row w:tr
      refElement = el.parentNode.parentNode.parentNode.parentNode
    }
  }

  refElement.parentNode.insertBefore(fakeElement, refElement)
}

function processClosingTag (doc, el, isVertical = false) {
  if (el.nodeName === 'w:t') {
    el.textContent = el.textContent.replace('{{/docxTable}}', '')
  }

  const fakeElement = doc.createElement('docxRemove')

  fakeElement.textContent = '{{/docxTable}}'

  let refElement

  if (el.nodeName !== 'w:t') {
    refElement = el
  } else {
    if (isVertical) {
      refElement = el.parentNode.parentNode.parentNode
    } else {
      refElement = el.parentNode.parentNode.parentNode.parentNode
    }
  }

  refElement.parentNode.insertBefore(fakeElement, refElement.nextSibling)
}

function getCellIndex (cellEl) {
  if (cellEl.nodeName !== 'w:tc') {
    throw new Error('Expected a table cell element during the processing')
  }

  let prevElements = 0

  let currentNode = cellEl.previousSibling

  while (
    currentNode != null &&
    currentNode.nodeName === 'w:tc'
  ) {
    prevElements += 1
    currentNode = currentNode.previousSibling
  }

  return prevElements
}

function getNextRowsUntilTableClose (rowEl) {
  if (rowEl.nodeName !== 'w:tr') {
    throw new Error('Expected a table row element during the processing')
  }

  let currentNode = rowEl.nextSibling
  let tableCloseNode
  const rows = []

  while (
    currentNode != null &&
    currentNode.nodeName === 'w:tr'
  ) {
    rows.push(currentNode)

    const cellNodes = nodeListToArray(currentNode.childNodes).filter((node) => node.nodeName === 'w:tc')

    for (const cellNode of cellNodes) {
      let textNodes = nodeListToArray(cellNode.getElementsByTagName('w:t'))

      // get text nodes of the current cell, we don't want text
      // nodes of nested tables
      textNodes = textNodes.filter((tNode) => {
        let current = tNode.parentNode

        while (current.nodeName !== 'w:tc') {
          current = current.parentNode
        }

        return current === cellNode
      })

      for (const tNode of textNodes) {
        if (tNode.textContent.includes('{{/docxTable')) {
          currentNode = null
          tableCloseNode = tNode
          break
        }
      }

      if (currentNode == null) {
        break
      }
    }

    if (currentNode != null) {
      currentNode = currentNode.nextSibling
    }
  }

  return [rows, tableCloseNode]
}
