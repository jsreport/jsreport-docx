/* eslint no-unused-vars: 1 */
/* eslint no-new-func: 0 */
/* *global __rootDirectory */
;(function (global) {
  const Handlebars = require('handlebars')

  global.docxPageBreak = function () {
    return new Handlebars.SafeString('')
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
        'docxImage helper requires src parameter to be set'
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

    return new Handlebars.SafeString('$docxImage' + Buffer.from(JSON.stringify({
      src: options.hash.src,
      width: options.hash.width,
      height: options.hash.height,
      usePlaceholderSize:
        options.hash.usePlaceholderSize === true ||
        options.hash.usePlaceholderSize === 'true'
    })).toString('base64') + '$')
  }

  global.docxCheckbox = function (options) {
    if (options.hash.value == null) {
      throw new Error('docxCheckbox helper requires value parameter')
    }

    options.hash.value = options.hash.value === 'true' || options.hash.value === true

    return new Handlebars.SafeString('$docxCheckbox' + Buffer.from(JSON.stringify(options.hash)).toString('base64') + '$')
  }

  global.docxCombobox = function (options) {
    return new Handlebars.SafeString('$docxCombobox' + Buffer.from(JSON.stringify(options.hash)).toString('base64') + '$')
  }

  global.docxChart = function (options) {
    if (options.hash.data == null) {
      throw new Error('docxChart helper requires data parameter to be set')
    }

    if (!Array.isArray(options.hash.data.labels) || options.hash.data.labels.length === 0) {
      throw new Error('docxChart helper requires data parameter with labels to be set, data.labels must be an array with items')
    }

    if (!Array.isArray(options.hash.data.datasets) || options.hash.data.datasets.length === 0) {
      throw new Error('docxChart helper requires data parameter with datasets to be set, data.datasets must be an array with items')
    }

    return new Handlebars.SafeString('$docxChart' + Buffer.from(JSON.stringify(options.hash)).toString('base64') + '$')
  }
})(this)
