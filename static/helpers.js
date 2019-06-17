/* eslint no-unused-vars: 1 */
/* eslint no-new-func: 0 */
/* *global __rootDirectory */
;(function (global) {
  const Handlebars = require('handlebars')
  global.docxList = function (data, options) {
    return Handlebars.helpers.each(data, options)
  }
  global.docxTable = function (data, options) {
    return Handlebars.helpers.each(data, options)
  }
  global.docxStyle = function (options) {
    return `<docxStyleStart textColor="${options.hash.textColor}" />${options.fn(this)}<docxStyleEnd/>`
  }

  global.docxImage = function (options) {
    if (!options.hash.src) {
      throw new Error('docxImage helper requires src parameter to be set')
    }

    if (!options.hash.src.startsWith('data:image/png;base64,') && !options.hash.src.startsWith('data:image/jpeg;base64,')) {
      throw new Error('docxImage helper requires src parameter to be valid data uri for png or jpeg image. got ' + options.hash.src)
    }

    return `<docxImage src="${options.hash.src}" />` + options.fn(this)
  }
})(this)
