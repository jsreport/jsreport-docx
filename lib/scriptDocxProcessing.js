const { DOMParser, XMLSerializer } = require('xmldom')
const decodeXML = require('unescape')
const { decompress, saveXmlsToOfficeFile } = require('jsreport-office')
const { sharedBuffer } = require('serializator')
const preprocess = require('./preprocess/preprocess.js')
const postprocess = require('./postprocess/postprocess.js')
const { contentIsXML } = require('./utils')

module.exports = async function scriptDocxProcessing (inputs, renderCallbackAsync) {
  const { docxTemplateContent, logger, options, outputPath } = inputs
  const docxTemplateBuf = sharedBuffer.createFrom(docxTemplateContent)

  try {
    const files = await decompress()(docxTemplateBuf)

    for (const f of files) {
      if (contentIsXML(f.data)) {
        f.doc = new DOMParser().parseFromString(f.data.toString())
        f.data = f.data.toString()
      }
    }

    await preprocess(files)

    const filesToRender = files.filter(f => contentIsXML(f.data))

    const contentToRender = filesToRender.map(f => {
      const xmlStr = new XMLSerializer().serializeToString(f.doc, undefined, (node) => {
        if (node.nodeType === 2 && node.nodeValue && node.nodeValue.includes('{{')) {
          const str = new XMLSerializer().serializeToString(node)
          return decodeXML(str)
        }

        return node
      })

      return xmlStr.replace(/<docxRemove>/g, '').replace(/<\/docxRemove>/g, '')
    }).join('$$$docxFile$$$')

    logger.debug('Starting child request to render docx dynamic parts')

    const { content: newContent } = await renderCallbackAsync({
      content: contentToRender
    })

    const contents = newContent.split('$$$docxFile$$$')

    for (let i = 0; i < filesToRender.length; i++) {
      filesToRender[i].data = contents[i]
      // don't parse the word/document.xml file, because after the templating engie execution
      // that documents can be a lot more bigger and parsing such big document is a performance
      // kill for the process
      if (filesToRender[i].path !== 'word/document.xml') {
        filesToRender[i].doc = new DOMParser().parseFromString(contents[i])
      } else {
        // we remove the .doc for the word/document.xml file to be clear that it should not be used
        // for any of postprocess steps, instead when dealing with that document we should executation search/replace
        // based on string and regexp.
        delete filesToRender[i].doc
      }
    }

    await postprocess(files, options)

    for (const f of files) {
      let shouldSerializeFromDoc = contentIsXML(f.data) && f.path !== 'word/document.xml'

      if (f.serializeFromDoc != null) {
        shouldSerializeFromDoc = f.serializeFromDoc === true
      }

      if (shouldSerializeFromDoc) {
        f.data = Buffer.from(new XMLSerializer().serializeToString(f.doc))
      }
    }

    await saveXmlsToOfficeFile({
      outputPath,
      files
    })

    logger.debug('docx successfully zipped')

    return {
      docxFilePath: outputPath
    }
  } catch (e) {
    return {
      error: {
        message: e.message,
        stack: e.stack
      }
    }
  }
}
