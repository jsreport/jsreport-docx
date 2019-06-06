const sizeOf = require('image-size')

function findNextImage (el) {
  if (el.nextSibling && el.nextSibling.getElementsByTagName('w:drawing')[0]) {
    return el.nextSibling.getElementsByTagName('w:drawing')[0]
  }

  if (el.nextSibling) {
    return findNextImage(el.nextSibling)
  }

  return findNextImage(el.parentNode)
}

module.exports = (doc, rels, newFiles) => {
  const elements = doc.getElementsByTagName('docxImage')
  const toRemove = []
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]
    toRemove.push(el)

    const imageSrc = el.getAttribute('src')
    const imageExtensions = imageSrc.split(';')[0].split('/')[1]
    const imageBuffer = Buffer.from(imageSrc.split(';')[1].substring('base64,'.length), 'base64')

    const relsCount = rels.getElementsByTagName('Relationship').length
    const id = relsCount + 1
    const relEl = rels.createElement('Relationship')
    relEl.setAttribute('Id', `rId${id}`)
    relEl.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    relEl.setAttribute('Target', `media/imageDocx${id}.${imageExtensions}`)

    newFiles.push({
      path: `word/media/imageDocx${id}.${imageExtensions}`,
      data: imageBuffer
    })

    rels.getElementsByTagName('Relationships')[0].appendChild(relEl)

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
