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

    const isValidDimensionUnit = (value) => {
      const regexp = /^(\d+(.\d+)?)(cm|px)$/
      return regexp.test(value)
    }

    if (options.hash.width != null && !isValidDimensionUnit(options.hash.width)) {
      throw new Error('docxImage helper requires width parameter to be valid number with unit (cm or px). got ' + options.hash.width)
    }

    if (options.hash.height != null && !isValidDimensionUnit(options.hash.height)) {
      throw new Error('docxImage helper requires height parameter to be valid number with unit (cm or px). got ' + options.hash.height)
    }

    return JSON.stringify({
      src: options.hash.src,
      width: options.hash.width,
      height: options.hash.height,
      usePlaceholderSize: options.hash.usePlaceholderSize === true || options.hash.usePlaceholderSize === 'true'
    })
  }
})(this)
