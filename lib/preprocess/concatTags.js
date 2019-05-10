
// excel splits strings like {{#each people}} into multiple xml nodes
// here we concat values from these splitted node and put it to one node
// so handlebars can correctly run
module.exports = (doc) => {
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
      if (value.endsWith('}}')) {
        elements[i].textContent = value + '$$$tag$$$'
        continue
      }
      startIndex = i
      tag = value
    }
  }

  for (const r of toRemove) {
    elements[r].parentNode.removeChild(elements[r])
  }
}
