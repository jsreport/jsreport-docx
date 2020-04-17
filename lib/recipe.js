const path = require('path')
const fs = require('fs')
const scriptCallbackRender = require('./scriptCallbackRender')
const { response } = require('jsreport-office')

module.exports = (reporter, definition) => async (req, res) => {
  if (!req.template.docx || (!req.template.docx.templateAsset && !req.template.docx.templateAssetShortid)) {
    throw reporter.createError(`docx requires template.docx.templateAsset or template.docx.templateAssetShortid to be set`, {
      statusCode: 400
    })
  }

  if (req.template.engine !== 'handlebars') {
    throw reporter.createError(`docx recipe can run only with handlebars`, {
      statusCode: 400
    })
  }

  let templateAsset = req.template.docx.templateAsset

  if (req.template.docx.templateAssetShortid) {
    templateAsset = await reporter.documentStore.collection('assets').findOne({ shortid: req.template.docx.templateAssetShortid }, req)

    if (!templateAsset) {
      throw reporter.createError(`Asset with shortid ${req.template.docx.templateAssetShortid} was not found`, {
        statusCode: 400
      })
    }
  } else {
    if (!Buffer.isBuffer(templateAsset.content)) {
      templateAsset.content = Buffer.from(templateAsset.content, templateAsset.encoding || 'utf8')
    }
  }

  reporter.logger.info('docx generation is starting', req)

  const { pathToFile: outputPath } = await reporter.writeTempFile((uuid) => `${uuid}.docx`, '')

  const result = await reporter.executeScript(
    {
      docxTemplateContent: templateAsset.content,
      options: {
        imageFetchParallelLimit: definition.options.imageFetchParallelLimit
      },
      outputPath
    },
    {
      execModulePath: path.join(__dirname, 'scriptDocxProcessing.js'),
      timeoutErrorMessage: 'Timeout during execution of docx recipe',
      callback: (params, cb) => scriptCallbackRender(reporter, req, params, cb)
    },
    req
  )

  if (result.logs) {
    result.logs.forEach(m => {
      reporter.logger[m.level](m.message, { ...req, timestamp: m.timestamp })
    })
  }

  if (result.error) {
    const error = new Error(result.error.message)
    error.stack = result.error.stack

    throw reporter.createError('Error while executing docx recipe', {
      original: error,
      weak: true
    })
  }

  reporter.logger.info('docx generation was finished', req)

  res.stream = fs.createReadStream(result.docxFilePath)

  await response({
    previewOptions: definition.options.preview,
    officeDocumentType: 'docx',
    stream: res.stream
  }, req, res)
}
