// the idea is that we find all w:t node between <docxStyleStart /> and <docxStyleEnd/>
// every such w:t element has in parent w:r node that has inside w:Pr where we add w:color element

function findNextText (p) {
  if (p.nextSibling && p.nextSibling.getElementsByTagName('w:t')[0]) {
    return p.nextSibling.getElementsByTagName('w:t')[0]
  }

  // with the recursion we go down the tree and searching for the next w:t

  if (p.nextSibling) {
    return findNextText(p.nextSibling)
  }

  return findNextText(p.parentNode)
}

function findRunsToStyle (styleStartText, result = []) {
  let run = styleStartText.parentNode
  while (run) {
    result.push(run)
    const t = run.getElementsByTagName('w:t')[0]
    if (t) {
      const styleEnd = t.getElementsByTagName('docxStyleEnd')[0]
      if (styleEnd) {
        return result
      }
    }
    run = run.nextSibling
  }

  return findRunsToStyle(findNextText(styleStartText.parentNode.parentNode), result)
}

module.exports = (doc) => {
  const elements = doc.getElementsByTagName('docxStyleStart')
  const toRemove = []
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i]
    toRemove.push(el)
    const runs = findRunsToStyle(el.parentNode)
    for (const wR of runs) {
      let wRpr = wR.getElementsByTagName('w:rPr')[0]
      if (!wRpr) {
        wRpr = doc.createElement('w:rPr')
        if (wR.childNodes.length === 0) {
          wR.appendChild(wRpr)
        } else {
          wR.insertBefore(wRpr, wR.getElementsByTagName('w:t')[0])
        }
      }
      let color = wRpr.getElementsByTagName('w:color')[0]
      if (!color) {
        color = doc.createElement('w:color')
        wRpr.appendChild(color)
      }
      color.setAttribute('w:val', el.getAttribute('textColor'))
      color.removeAttribute('w:themeColor')
    }
    const runWithSyleEnd = runs.find((r) => r.getElementsByTagName('w:t')[0] && r.getElementsByTagName('w:t')[0].getElementsByTagName('docxStyleEnd')[0])
    toRemove.push(runWithSyleEnd.getElementsByTagName('w:t')[0].getElementsByTagName('docxStyleEnd')[0])
  }

  // remove docxStyleStart and docxStyleEnd
  for (const el of toRemove) {
    el.parentNode.removeChild(el)
  }
}
