const recipe = require('./recipe')

module.exports = (reporter, definition) => {
  reporter.extensionsManager.recipes.push({
    name: 'docx',
    execute: recipe(reporter, definition)
  })

  reporter.documentStore.registerComplexType('DocxType', {
    templateAssetShortid: { type: 'Edm.String' }
  })

  reporter.documentStore.model.entityTypes['TemplateType'].docx = { type: 'jsreport.DocxType' }

  reporter.beforeRenderListeners.insert({ before: 'templates' }, 'docx', (req) => {
    if (req.template && req.template.recipe === 'docx' && !req.template.name && !req.template.shortid && !req.template.content) {
      // templates extension otherwise complains that the template is empty
      // but that is fine for this recipe
      req.template.content = 'docx placeholder'
    }
  })

  reporter.initializeListeners.add('docx', () => {
    if (reporter.express) {
      reporter.express.exposeOptionsToApi(definition.name, {
        previewInWordOnline: definition.options.previewInWordOnline,
        showWordOnlineWarning: definition.options.showWordOnlineWarning
      })
    }
  })
}
