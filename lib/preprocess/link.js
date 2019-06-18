const { nodeListToArray } = require('../utils')

// hyperlink value is url encoded so developer cannot use handlebars there
// we url decode this cases before handlebars runs
// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%7b%7burl%7d%7d" TargetMode="External"/>
module.exports = (files) => {
  const docRels = files.find(f => f.path === 'word/_rels/document.xml.rels').doc

  const rels = nodeListToArray(docRels.getElementsByTagName('Relationships')[0].getElementsByTagName('Relationship'))

  const hyperlinkRels = rels.filter(r => r.getAttribute('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink')
  for (const hyperlink of hyperlinkRels) {
    const decodedTarget = decodeURIComponent(hyperlink.getAttribute('Target'))
    if (decodedTarget.includes('{{')) {
      hyperlink.setAttribute('Target', decodedTarget)
    }
  }
}
