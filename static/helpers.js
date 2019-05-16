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
  global.docxStyle = function (context, options) {
    return options.fn(context)
  }
  /* return Handlebars.SafeString(
      `<docxList><data>${JSON.stringify(data)}</data></docxList>${Handlebars.Utils.escapeExpression(options.fn())}<docxListEnd/>`
    ) */
})(this)
