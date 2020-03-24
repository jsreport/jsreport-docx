const concatTextNodes = require('./_concatTextNodes')

// excel splits string like {{docxChart ...}} into multiple xml nodes
// here we concat values from these splitted node and put it to one node
// so handlebars can correctly run
module.exports = (files) => {
  for (const f of files.filter(f => f.path.startsWith('word/charts/') && f.path.endsWith('.xml'))) {
    const doc = f.doc
    const elements = doc.getElementsByTagName('a:t')
    concatTextNodes(elements, { removeParent: true })
  }
}
