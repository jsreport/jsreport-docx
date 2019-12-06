
module.exports = async function render (reporter, req, { content, logs }, cb) {
  try {
    if (logs) {
      // we handle logs here in callback in order to mantain correct order of logs
      // between render callback calls
      logs.forEach((m) => {
        reporter.logger[m.level](m.message, { ...req, timestamp: m.timestamp })
      })
    }

    // do an anonymous render
    const template = {
      content,
      engine: req.template.engine,
      recipe: 'html',
      helpers: req.template.helpers
    }

    const result = await reporter.render({ template }, req)

    cb(null, {
      content: result.content.toString()
    })
  } catch (e) {
    cb(e)
  }
}
