const path = require('path')
module.exports = async (reporter, {
  pathToEngine,
  req,
  content,
  helpers,
  data
}) => {
  const pathToEngineScript = path.join(path.dirname(require.resolve('jsreport-core')), 'lib', 'render', 'engineScript.js')

  const engineRes = await reporter.executeScript({
    template: {
      content, helpers
    },
    data: data,
    engine: pathToEngine,
    safeSandboxPath: reporter.options.templatingEngines.safeSandboxPath,
    appDirectory: reporter.options.appDirectory,
    rootDirectory: reporter.options.rootDirectory,
    parentModuleDirectory: reporter.options.parentModuleDirectory,
    templatingEngines: reporter.options.templatingEngines
  }, {
    execModulePath: pathToEngineScript,
    timeoutErrorMessage: 'Timeout during execution of templating engine'
  }, req)

  engineRes.logs.forEach(function (m) {
    reporter.logger[m.level](m.message, { req, timestamp: m.timestamp })
  })

  return engineRes.content.toString()
}
