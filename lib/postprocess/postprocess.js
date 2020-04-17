const style = require('./style')
const image = require('./image')
const chart = require('./chart')
const link = require('./link')
const form = require('./form')
const pageBreak = require('./pageBreak')

module.exports = async (files, options) => {
  await pageBreak(files)
  style(files)
  await image(files, options)
  await chart(files)
  link(files)
  form(files)
}
