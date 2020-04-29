
module.exports = async function render (reporter, req, { content }) {
  // do an anonymous render
  const template = {
    content,
    engine: req.template.engine,
    recipe: 'html',
    helpers: req.template.helpers
  }

  const result = await reporter.render({ template }, req)

  return {
    content: result.content.toString()
  }
}
