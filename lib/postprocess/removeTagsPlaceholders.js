
module.exports = (files) => {
  const doc = files.find(f => f.path === 'word/document.xml').doc
  const toRemove = []
  const elements = doc.getElementsByTagName('w:t')
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]
    if (el.textContent === '$$$tag$$$') {
      toRemove.push(el)
      continue
    }

    el.textContent = el.textContent.replace(/\$\$\$tag\$\$\$/g, '')
  }

  for (const el of toRemove) {
    el.parentNode.removeChild(el)
  }
}
