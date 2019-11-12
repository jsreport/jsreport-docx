const style = require('./style')
const image = require('./image')
const link = require('./link')
const form = require('./form')

module.exports = (files) => {
  style(files)
  image(files)
  link(files)
  form(files)
}
