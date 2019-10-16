
// excel splits strings like {{#each people}} into multiple xml nodes
// here we concat values from these splitted node and put it to one node
// so handlebars can correctly run
module.exports = (files) => {
  for (const f of files.filter(f => f.path.endsWith('.xml'))) {
    const doc = f.doc
    const elements = doc.getElementsByTagName('w:t')

    const toRemove = []
    const lastIndex = elements.length - 1
    let startIndex = -1
    let tag = ''

    for (let i = 0; i < elements.length; i++) {
      const value = elements[i].textContent
      let concatenating = startIndex !== -1
      let toEvaluate = value
      let validSiblings = false

      if (concatenating) {
        if (elements[i].parentNode.previousSibling.localName === 'r') {
          validSiblings = elements[i].parentNode.previousSibling === elements[i - 1].parentNode
        } else {
          // ignore w:proofErr, w:bookmarkStart and other similar self-closed siblings tags
          let currentNode = elements[i].parentNode
          const previousSiblings = []

          while (currentNode && currentNode.previousSibling != null) {
            if (currentNode.previousSibling.localName === 'r') {
              previousSiblings.push(currentNode.previousSibling)
            }

            currentNode = currentNode.previousSibling
          }

          validSiblings = previousSiblings.some((s) => s === elements[i - 1].parentNode)
        }
      }

      // checking that nodes are valid siblings for the concat to be valid, if they are not siblings stop
      // concatenation at current index, this prevents concatenating text with bad syntax with lists
      if (concatenating && !validSiblings) {
        elements[startIndex].textContent = tag
        concatenating = false
        tag = ''
        startIndex = -1
      }

      if (concatenating) {
        toEvaluate = tag + value

        if (!value && elements[i].getAttribute('xml:space') === 'preserve') {
          toEvaluate += ' '
        }
      }

      const openTags = matchRegExp(toEvaluate, '{', 'g')
      const closingTags = matchRegExp(toEvaluate, '}', 'g')
      let shouldCheckTag = false

      if (
        (openTags.length > 0 && openTags.length === closingTags.length) ||
        // if it is incomplete and we are already on last node
        (concatenating && i === lastIndex)
      ) {
        tag = ''
        shouldCheckTag = true

        if (concatenating) {
          elements[startIndex].textContent = toEvaluate
          startIndex = -1
          toRemove.push(i)
        }
      } else if (openTags.length > 0 && openTags.length !== closingTags.length) {
        tag = toEvaluate

        if (concatenating) {
          toRemove.push(i)
        } else {
          startIndex = i
        }
      }

      if (shouldCheckTag && toEvaluate.endsWith('}}')) {
        elements[i].textContent = toEvaluate + '$$$tag$$$'
      }
    }

    for (const r of toRemove) {
      elements[r].parentNode.removeChild(elements[r])
    }
  }
}

function matchRegExp (str, pattern, flags) {
  let f = flags || ''
  let r = new RegExp(pattern, 'g' + f.replace(/g/g, ''))
  let a = []
  let m

  // eslint-disable-next-line no-cond-assign
  while (m = r.exec(str)) {
    a.push({
      index: m.index,
      offset: r.lastIndex
    })
  }

  return a
}
