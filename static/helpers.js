/* eslint no-unused-vars: 1 */
/* eslint no-new-func: 0 */
/* *global __rootDirectory */
;(function (global) {
  const Handlebars = require('handlebars')

  global.docxPageBreak = function () {
    return new Handlebars.SafeString('$docxPageBreak')
  }

  global.docxList = function (data, options) {
    return Handlebars.helpers.each(data, options)
  }
  global.docxTable = function (data, options) {
    return Handlebars.helpers.each(data, options)
  }
  global.docxStyle = function (options) {
    return new Handlebars.SafeString(
      `<docxStyle id="${options.hash.id}" textColor="${options.hash.textColor}" />`
    )
  }

  global.docxImage = function (options) {
    if (!options.hash.src) {
      throw new Error(
        'docxImage helper requires url parameter to be set'
      )
    }

    if (
      !options.hash.src.startsWith('data:image/png;base64,') &&
      !options.hash.src.startsWith('data:image/jpeg;base64,') &&
      !options.hash.src.startsWith('http://') &&
      !options.hash.src.startsWith('https://')
    ) {
      throw new Error(
        'docxImage helper requires src parameter to be valid data uri for png or jpeg image or a valid url. Got ' +
          options.hash.src
      )
    }

    const isValidDimensionUnit = value => {
      const regexp = /^(\d+(.\d+)?)(cm|px)$/
      return regexp.test(value)
    }

    if (
      options.hash.width != null &&
      !isValidDimensionUnit(options.hash.width)
    ) {
      throw new Error(
        'docxImage helper requires width parameter to be valid number with unit (cm or px). got ' +
          options.hash.width
      )
    }

    if (
      options.hash.height != null &&
      !isValidDimensionUnit(options.hash.height)
    ) {
      throw new Error(
        'docxImage helper requires height parameter to be valid number with unit (cm or px). got ' +
          options.hash.height
      )
    }

    return JSON.stringify({
      src: options.hash.src,
      width: options.hash.width,
      height: options.hash.height,
      usePlaceholderSize:
        options.hash.usePlaceholderSize === true ||
        options.hash.usePlaceholderSize === 'true'
    })
  }

  global.docxForm = function (options) {
    const hash = options.hash || {}
    return JSON.stringify({ ...hash })
  }
})(this)
