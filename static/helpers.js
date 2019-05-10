/* eslint no-unused-vars: 1 */
/* eslint no-new-func: 0 */
/* *global __rootDirectory */
;(function (global) {
  global.docxList = function (data, options) {
    let results = '<jsreport:list>'
    for (const e of data) {
      results += `<jsreport:item>${options.fn(e)}</jsreport:item>`
    }

    return results + '</jsreport:list>'
  }
})(this)
