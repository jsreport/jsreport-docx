const { DOMParser, XMLSerializer } = require('xmldom')

module.exports = (content) => {
  const doc = new DOMParser().parseFromString(content)
  const elements = doc.getElementsByTagName('w:t')

  const toRemove = []
  let startIndex = -1
  let tag = ''

  for (let i = 0; i < elements.length; i++) {
    const value = elements[i].textContent

    if (startIndex !== -1) {
      tag += value

      if (!value && elements[i].getAttribute('xml:space') === 'preserve') {
        tag += ' '
      }

      if (tag.endsWith('}}')) {
        elements[startIndex].textContent = tag
        startIndex = -1
      }

      toRemove.push(i)
      continue
    }

    const indexStart = value.indexOf('{')
    if (indexStart !== -1) {
      startIndex = i
      tag = value
    }
  }

  for (const r of toRemove) {
    elements[r].parentNode.removeChild(elements[r])
  }

  return new XMLSerializer().serializeToString(doc)
}
