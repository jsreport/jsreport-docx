
module.exports = (files) => {
  for (const f of files.filter(f => f.path.endsWith('.xml'))) {
    const doc = f.doc
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
}
