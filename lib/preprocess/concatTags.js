const concatTextNodes = require('./_concatTextNodes')

// excel splits strings like {{#each people}} into multiple xml nodes
// here we concat values from these splitted node and put it to one node
// so handlebars can correctly run
module.exports = (files) => {
  for (const f of files.filter(f => f.path.endsWith('.xml'))) {
    const doc = f.doc
    const elements = doc.getElementsByTagName('w:t')
    concatTextNodes(elements)
  }
}
