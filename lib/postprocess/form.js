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

    const formStateEl = controlEl.parentNode.getElementsByTagName('docxFormState')[0]

    if (!formStateEl) {
      continue
    }

    const formState = JSON.parse(formStateEl.textContent)

    if (formState.value == null) {
      continue
    }

    let controlValue

    if (tagName === 'w:checkBox') {
      controlValue = formState.value === true || formState.value === 'true'
      controlEl.getElementsByTagName('w:default')[0].setAttribute('w:val', controlValue ? '1' : '0')
    } else if (tagName === 'w:ddList') {
      controlValue = Array.isArray(formState.value) ? formState.value : []

      const childNodes = nodeListToArray(controlEl.childNodes)

      childNodes.forEach((node) => {
        controlEl.removeChild(node)
      })

      controlValue.forEach((item) => {
        const node = doc.createElement('w:listEntry')
        node.setAttribute('w:val', item)
        controlEl.appendChild(node)
      })
    }

    formStateEl.parentNode.removeChild(formStateEl)
  }
}
