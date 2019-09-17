const office = require('jsreport-office')

module.exports = {
  'name': 'docx',
  'main': 'lib/docx.js',
  'optionsSchema': office.extendSchema('docx', {
    type: 'object',
    properties: {
      beta: {
        type: 'object',
        properties: {
          showWarning: { type: 'boolean', default: true }
        }
      }
    }
  }),
  'dependencies': []
}
