const { nodeListToArray } = require('../utils')

module.exports = (files) => {
  const f = files.find(f => f.path === 'word/document.xml')
  const doc = f.doc

  processControlElements(doc, 'w:checkBox')
  processControlElements(doc, 'w:ddList')
}

function processControlElements (doc, tagName) {
  const controls = doc.getElementsByTagName(tagName)

  for (let i = 0; i < controls.length; i++) {
    const controlEl = controls[i]

    const statusNode = nodeListToArray(controlEl.parentNode.getElementsByTagName('w:statusText')).filter((el) => {
      return el.getAttribute('w:type') === 'text'
    })[0]

    if (!statusNode) {
      continue
    }

    let statusText = statusNode.getAttribute('w:val')

    const results = matchRecursiveRegExp(statusText, '{', '}', 'g')
    const dynamic = []
    let removed = 0

    results.forEach((info, idx) => {
      if (!info.match.startsWith('{docxForm')) {
        return
      }

      dynamic.push(`{${info.match}}`)

      const currentOffset = info.offset - removed
      removed += info.match.length + 2 // 2 because the "{", "}" delimiters of the match

      statusText = `${statusText.slice(0, currentOffset - 1)}${statusText.slice(currentOffset + info.match.length + 1)}`
    })

    if (dynamic.length === 0) {
      continue
    }

    // set new text without the dynamic parts
    statusNode.setAttribute('w:val', statusText)

    const formStateEl = doc.createElement('docxFormState')

    formStateEl.textContent = dynamic[dynamic.length - 1]
    statusNode.parentNode.insertBefore(formStateEl, statusNode.nextSibling)
  }
}

// taken from: http://blog.stevenlevithan.com/archives/javascript-match-recursive-regexp
function matchRecursiveRegExp (str, left, right, flags) {
  let f = flags || ''
  let g = f.indexOf('g') > -1
  let x = new RegExp(left + '|' + right, 'g' + f.replace(/g/g, ''))
  let l = new RegExp(left, f.replace(/g/g, ''))
  let a = []
  let t
  let s
  let m

  do {
    t = 0

    // eslint-disable-next-line no-cond-assign
    while (m = x.exec(str)) {
      if (l.test(m[0])) {
        if (!t++) s = x.lastIndex
      } else if (t) {
        if (!--t) {
          const match = str.slice(s, m.index)
          a.push({
            offset: s,
            match
          })

          if (!g) return a
        }
      }
    }
  } while (t && (x.lastIndex = s))

  return a
}
