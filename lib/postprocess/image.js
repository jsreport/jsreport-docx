const sizeOf = require('image-size')
const { nodeListToArray, pxToEMU, cmToEMU } = require('../utils')
const { DOMParser, XMLSerializer } = require('xmldom')

const getDimension = (value) => {
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

module.exports = (files) => {
  const contentTypesFile = files.find(f => f.path === '[Content_Types].xml')
  const types = contentTypesFile.doc.getElementsByTagName('Types')[0]

  let pngDefault = nodeListToArray(types.getElementsByTagName('Default')).find(d => d.getAttribute('Extension') === 'png')

  if (!pngDefault) {
    const defaultPng = contentTypesFile.doc.createElement('Default')
    defaultPng.setAttribute('Extension', 'png')
    defaultPng.setAttribute('ContentType', 'image/png')
    types.appendChild(defaultPng)
  }

  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc

  const documentFile = files.find(f => f.path === 'word/document.xml')
  documentFile.data = documentFile.data.toString().replace(/<w:drawing>.*?(?=<\/w:drawing>)<\/w:drawing>/g, (val) => {
    const elDrawing = new DOMParser().parseFromString(val)
    const elLinkClick = elDrawing.getElementsByTagName('a:hlinkClick')[0]

    if (!elLinkClick) {
      return val
    }

    const imageConfig = JSON.parse(decodeURIComponent(elLinkClick.getAttribute('tooltip')))

    const imageSrc = imageConfig.src
    const imageExtensions = imageSrc.split(';')[0].split('/')[1]
    const imageBuffer = Buffer.from(imageSrc.split(';')[1].substring('base64,'.length), 'base64')

    const relsElements = nodeListToArray(relsDoc.getElementsByTagName('Relationship'))
    const relsCount = relsElements.length
    const id = relsCount + 1
    const relEl = relsDoc.createElement('Relationship')
    relEl.setAttribute('Id', `rId${id}`)
    relEl.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    relEl.setAttribute('Target', `media/imageDocx${id}.${imageExtensions}`)

    files.push({
      path: `word/media/imageDocx${id}.${imageExtensions}`,
      data: imageBuffer
    })

    relsDoc.getElementsByTagName('Relationships')[0].appendChild(relEl)

    const relPlaceholder = elDrawing.getElementsByTagName('a:blip')[0]
    const wpExtendEl = elDrawing.getElementsByTagName('wp:extent')[0]

    let imageWidthEMU
    let imageHeightEMU

    if (imageConfig.width != null || imageConfig.height != null) {
      const imageDimension = sizeOf(imageBuffer)
      const targetWidth = getDimension(imageConfig.width)
      const targetHeight = getDimension(imageConfig.height)

      if (targetWidth) {
        imageWidthEMU = targetWidth.unit === 'cm' ? cmToEMU(targetWidth.value) : pxToEMU(targetWidth.value)
      }

      if (targetHeight) {
        imageHeightEMU = targetHeight.unit === 'cm' ? cmToEMU(targetHeight.value) : pxToEMU(targetHeight.value)
      }

      if (imageWidthEMU != null && imageHeightEMU == null) {
      // adjust height based on aspect ratio of image
        imageHeightEMU = Math.round(imageWidthEMU * (pxToEMU(imageDimension.height) / pxToEMU(imageDimension.width)))
      } else if (imageHeightEMU != null && imageWidthEMU == null) {
      // adjust width based on aspect ratio of image
        imageWidthEMU = Math.round(imageHeightEMU * (pxToEMU(imageDimension.width) / pxToEMU(imageDimension.height)))
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

    relPlaceholder.setAttribute('r:embed', `rId${id}`)

    wpExtendEl.setAttribute('cx', imageWidthEMU)
    wpExtendEl.setAttribute('cy', imageHeightEMU)
    const aExtEl = elDrawing.getElementsByTagName('a:xfrm')[0].getElementsByTagName('a:ext')[0]
    aExtEl.setAttribute('cx', imageWidthEMU)
    aExtEl.setAttribute('cy', imageHeightEMU)

    elLinkClick.parentNode.removeChild(elLinkClick)

    return new XMLSerializer().serializeToString(elDrawing)
  })
}
