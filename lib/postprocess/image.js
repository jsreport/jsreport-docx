const sizeOf = require('image-size')
const { nodeListToArray } = require('../utils')

function findNextImage (el) {
  if (el.nextSibling && el.nextSibling.getElementsByTagName('w:drawing')[0]) {
    return el.nextSibling.getElementsByTagName('w:drawing')[0]
  }

  if (el.nextSibling) {
    return findNextImage(el.nextSibling)
  }

  return findNextImage(el.parentNode)
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

  const doc = files.find(f => f.path === 'word/document.xml').doc

  const relsDoc = files.find(f => f.path === 'word/_rels/document.xml.rels').doc
  const elements = doc.getElementsByTagName('docxImage')
  const toRemove = []
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]
    toRemove.push(el)

    const imageSrc = el.getAttribute('src')
    const imageExtensions = imageSrc.split(';')[0].split('/')[1]
    const imageBuffer = Buffer.from(imageSrc.split(';')[1].substring('base64,'.length), 'base64')

    const relsCount = relsDoc.getElementsByTagName('Relationship').length
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

    const drawing = findNextImage(el)
    const relPlaceholder = drawing.getElementsByTagName('a:blip')[0]
    relPlaceholder.setAttribute('r:embed', `rId${id}`)

    const imageDimension = sizeOf(imageBuffer)
    const imageWidthEMU = Math.round(imageDimension.width * 914400 / 96)
    const imageHeightEMU = Math.round(imageDimension.height * 914400 / 96)
    const wpExtendEl = drawing.getElementsByTagName('wp:extent')[0]
    wpExtendEl.setAttribute('cx', imageWidthEMU)
    wpExtendEl.setAttribute('cy', imageHeightEMU)
    const aExtEl = drawing.getElementsByTagName('a:xfrm')[0].getElementsByTagName('a:ext')[0]
    aExtEl.setAttribute('cx', imageWidthEMU)
    aExtEl.setAttribute('cy', imageHeightEMU)
  }

  for (const el of toRemove) {
    el.parentNode.removeChild(el)
  }
}
