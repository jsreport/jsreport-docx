const schema = {
  type: 'object',
  properties: {
    previewInOfficeOnline: { type: 'boolean' },
    publicUriForPreview: { type: 'string' },
    showOfficeOnlineWarning: { type: 'boolean', default: true }
  }
}
module.exports = {
  'name': 'docx',
  'main': 'lib/docx.js',
  'optionsSchema': {
    extensions: {
      docx: { ...schema }
    }
  },
  'dependencies': []
}
