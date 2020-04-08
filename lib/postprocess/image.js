const sizeOf = require('image-size')
const axios = require('axios')

const recursiveStringReplaceAsync = require('../recursiveStringReplaceAsync')
const { nodeListToArray, serializeXml, pxToEMU, cmToEMU, getNewRelIdFromBaseId, getNewRelId } = require('../utils')
const { DOMParser } = require('xmldom')

const getDimension = value => {
  const regexp = /^(\d+(.\d+)?)(cm|px)$/
  const match = regexp.exec(value)

  if (match) {
    return {
      value: parseFloat(match[1]),
      unit: match[3]
    }
  }

  return null
}

module.exports = async files => {
  const contentTypesFile = files.find(f => f.path === '[Content_Types].xml')
  const types = contentTypesFile.doc.getElementsByTagName('Types')[0]

  let pngDefault = nodeListToArray(types.getElementsByTagName('Default')).find(
    d => d.getAttribute('Extension') === 'png'
  )

  if (!pngDefault) {
    const defaultPng = contentTypesFile.doc.createElement('Default')
    defaultPng.setAttribute('Extension', 'png')
    defaultPng.setAttribute('ContentType', 'image/png')
    types.appendChild(defaultPng)
  }

  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const relsEl = relsDoc.getElementsByTagName('Relationships')[0]
  const documentFile = files.find(f => f.path === 'word/document.xml')
  const newRelIdCounterMap = new Map()

  documentFile.data = await recursiveStringReplaceAsync(
    documentFile.data.toString(),
    '<w:drawing>',
    '</w:drawing>',
    'g',
    async (val, content, hasNestedMatch) => {
      if (hasNestedMatch) {
        return val
      }

      // no need to pass xml namespaces here because the nodes there are just used for reads,
      // and are not inserted (re-used) somewhere else
      const elDrawing = new DOMParser().parseFromString(val)
      const isImg = elDrawing.getElementsByTagName('pic:pic').length > 0

      if (!isImg) {
        return val
      }

      const elLinkClicks = elDrawing.getElementsByTagName('a:hlinkClick')
      const elLinkClick = elLinkClicks[0]

      if (!elLinkClick) {
        return val
      }

      if (elDrawing.documentElement.firstChild.nodeName === 'Relationship') {
        const hyperlinkRelEl = elDrawing.documentElement.firstChild
        const newHyperlinkRelId = getNewRelIdFromBaseId(relsDoc, newRelIdCounterMap, hyperlinkRelEl.getAttribute('Id'))

        if (hyperlinkRelEl.getAttribute('Id') === newHyperlinkRelId) {
          // if we get the same id it means that we should replace old rel node
          const oldRelEl = nodeListToArray(relsEl.getElementsByTagName('Relationship')).find((el) => {
            return el.getAttribute('Id') === newHyperlinkRelId
          })

          oldRelEl.parentNode.removeChild(oldRelEl)
        }

        hyperlinkRelEl.setAttribute('Id', newHyperlinkRelId)
        hyperlinkRelEl.parentNode.removeChild(hyperlinkRelEl)

        relsEl.appendChild(hyperlinkRelEl)

        elLinkClick.setAttribute('r:id', newHyperlinkRelId)

        for (let i = 1; i < elLinkClicks.length; i++) {
          const elLinkClick = elLinkClicks[i]
          elLinkClick.setAttribute('r:id', newHyperlinkRelId)
        }
      }

      const tooltip = elLinkClick.getAttribute('tooltip')

      if (tooltip == null || !tooltip.includes('$docxImage')) {
        return val
      }

      const match = tooltip.match(/\$docxImage([^$]*)\$/)

      elLinkClick.setAttribute('tooltip', tooltip.replace(match[0], ''))

      const imageConfig = JSON.parse(Buffer.from(match[1], 'base64').toString())

      // somehow there are duplicated hlinkclick els produced by word, we need to clean them up
      for (let i = 1; i < elLinkClicks.length; i++) {
        const elLinkClick = elLinkClicks[i]
        const match = tooltip.match(/\$docxImage([^$]*)\$/)
        elLinkClick.setAttribute('tooltip', tooltip.replace(match[0], ''))
      }

      let imageBuffer
      let imageExtensions

      if (imageConfig.src && imageConfig.src.startsWith('data:')) {
        const imageSrc = imageConfig.src
        imageExtensions = imageSrc.split(';')[0].split('/')[1]
        imageBuffer = Buffer.from(
          imageSrc.split(';')[1].substring('base64,'.length),
          'base64'
        )
      } else {
        const response = await axios({
          url: imageConfig.src,
          responseType: 'arraybuffer',
          method: 'GET'
        })
        const contentType = response.headers['content-type'] || response.headers['Content-Type']
        imageExtensions = contentType.split('/')[1]
        imageBuffer = Buffer.from(response.data)
      }

      const newImageRelId = getNewRelId(relsDoc)

      const relEl = relsDoc.createElement('Relationship')

      relEl.setAttribute('Id', newImageRelId)

      relEl.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
      )

      relEl.setAttribute('Target', `media/imageDocx${newImageRelId}.${imageExtensions}`)

      files.push({
        path: `word/media/imageDocx${newImageRelId}.${imageExtensions}`,
        data: imageBuffer
      })

      relsEl.appendChild(relEl)

      const relPlaceholder = elDrawing.getElementsByTagName('a:blip')[0]
      const wpExtendEl = elDrawing.getElementsByTagName('wp:extent')[0]

      let imageWidthEMU
      let imageHeightEMU

      if (imageConfig.width != null || imageConfig.height != null) {
        const imageDimension = sizeOf(imageBuffer)
        const targetWidth = getDimension(imageConfig.width)
        const targetHeight = getDimension(imageConfig.height)

        if (targetWidth) {
          imageWidthEMU =
            targetWidth.unit === 'cm'
              ? cmToEMU(targetWidth.value)
              : pxToEMU(targetWidth.value)
        }

        if (targetHeight) {
          imageHeightEMU =
            targetHeight.unit === 'cm'
              ? cmToEMU(targetHeight.value)
              : pxToEMU(targetHeight.value)
        }

        if (imageWidthEMU != null && imageHeightEMU == null) {
          // adjust height based on aspect ratio of image
          imageHeightEMU = Math.round(
            imageWidthEMU *
              (pxToEMU(imageDimension.height) / pxToEMU(imageDimension.width))
          )
        } else if (imageHeightEMU != null && imageWidthEMU == null) {
          // adjust width based on aspect ratio of image
          imageWidthEMU = Math.round(
            imageHeightEMU *
              (pxToEMU(imageDimension.width) / pxToEMU(imageDimension.height))
          )
        }
      } else if (imageConfig.usePlaceholderSize) {
        // taking existing size defined in word
        imageWidthEMU = parseFloat(wpExtendEl.getAttribute('cx'))
        imageHeightEMU = parseFloat(wpExtendEl.getAttribute('cy'))
      } else {
        const imageDimension = sizeOf(imageBuffer)
        imageWidthEMU = pxToEMU(imageDimension.width)
        imageHeightEMU = pxToEMU(imageDimension.height)
      }

      relPlaceholder.setAttribute('r:embed', newImageRelId)

      wpExtendEl.setAttribute('cx', imageWidthEMU)
      wpExtendEl.setAttribute('cy', imageHeightEMU)
      const aExtEl = elDrawing
        .getElementsByTagName('a:xfrm')[0]
        .getElementsByTagName('a:ext')[0]
      aExtEl.setAttribute('cx', imageWidthEMU)
      aExtEl.setAttribute('cy', imageHeightEMU)

      return serializeXml(elDrawing)
    }
  )
}
