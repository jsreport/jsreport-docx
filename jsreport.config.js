const office = require('jsreport-office')

module.exports = {
  'name': 'docx',
  'main': 'lib/docx.js',
  'optionsSchema': office.extendSchema('docx', {}),
  'dependencies': []
}
