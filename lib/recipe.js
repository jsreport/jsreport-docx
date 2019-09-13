const extend = require('node.extend.without.arrays')
const { DOMParser, XMLSerializer } = require('xmldom')
const { decompress, response, serializeOfficeXmls } = require('jsreport-office')
const preprocess = require('./preprocess/preprocess.js')
const postprocess = require('./postprocess/postprocess.js')

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

  const files = await decompress()(templateAsset.content)

  for (const f of files) {
    if (!f.path.includes('/media')) {
      f.doc = new DOMParser().parseFromString(f.data.toString())
      f.data = f.data.toString()
    }
  }

  await preprocess(files)

  const filesToRender = files.filter(f => !f.path.includes('/media'))
  const contentToRender = filesToRender
    .map(f => new XMLSerializer().serializeToString(f.doc).replace(/<docxRemove>/g, '').replace(/<\/docxRemove>/g, ''))
    .join('$$$docxFile$$$')

  reporter.logger.debug(`Starting child request to render docx dynamic parts`, req)

  // delete _id, shortid, name to do an anonymous render
  const template = extend(true, {}, req.template, {
    _id: null,
    shortid: null,
    name: null,
    content: contentToRender,
    recipe: 'html'
  })

  const renderResult = await reporter.render({ template }, req)
  const contents = renderResult.content.toString().split('$$$docxFile$$$')
  for (let i = 0; i < filesToRender.length; i++) {
    filesToRender[i].data = contents[i]
    filesToRender[i].doc = new DOMParser().parseFromString(contents[i])
  }

  await postprocess(files)

  for (const f of files) {
    if (!f.path.includes('/media')) {
      f.data = Buffer.from(new XMLSerializer().serializeToString(f.doc))
    }
  }

  await serializeOfficeXmls({ reporter, files, officeDocumentType: 'docx' }, req, res)

  await response({
    previewOptions: definition.options.preview,
    officeDocumentType: 'docx',
    stream: res.stream
  }, req, res)
}
