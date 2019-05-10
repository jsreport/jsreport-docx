require('should')
const jsreport = require('jsreport-core')
const fs = require('fs')
const path = require('path')
const util = require('util')
const textract = util.promisify(require('textract').fromBufferWithName)

describe('docx', () => {
  let reporter

  beforeEach(() => {
    reporter = jsreport({
      templatingEngines: {
        strategy: 'in-process'
      }
    }).use(require('../')())
      .use(require('jsreport-handlebars')())
      .use(require('jsreport-templates')())
      .use(require('jsreport-assets')())
    return reporter.init()
  })

  afterEach(() => reporter.close())

  it('variable-replace', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace.docx'))
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Hello world John')
  })

  it('variable-replace', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'list.docx'))
          }
        }
      },
      data: {
        people: [{
          name: 'Jan'
        }, {
          name: 'Boris'
        }, {
          name: 'Pavel'
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Jan')
    text.should.containEql('Boris')
  })

  it('loop', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'loop.docx'))
          }
        }
      },
      data: {
        chapters: [{
          title: 'Chapter 1',
          text: 'This is the first chapter'
        }, {
          title: 'Chapter 2',
          text: 'This is the second chapter'
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Chapter 1')
    text.should.containEql('This is the first chapter')
    text.should.containEql('Chapter 2')
    text.should.containEql('This is the second chapter')
  })

  it('complex', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'complex.docx'))
          }
        }
      },
      data: {
        name: 'Jan Blaha',
        email: 'jan.blaha@jsreport.net',
        phone: '+420777271254',
        description: `I am software developer, software architect and consultant with over 8 years of professional
        experience working on projects for cross domain market leaders. My experience covers custom
        projects for big costumers in the banking or electricity domain as well as cloud based SaaS startups.`,
        experiences: [{
          title: '.NET Developer',
          company: 'Unicorn',
          from: '1.1.2010',
          to: '15.5.2012'
        }, {
          title: 'Solution Architect',
          company: 'Simplias',
          from: '15.5.2012',
          to: 'now'
        }],
        skills: [{
          title: 'The worst developer ever'
        }, {
          title: `Don't need to write semicolons`
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Jan Blaha')
  })

  it('should be able to reference stored asset', async () => {
    await reporter.documentStore.collection('assets').insert({
      name: 'variable-replace.docx',
      shortid: 'template',
      content: fs.readFileSync(path.join(__dirname, 'variable-replace.docx'))
    })
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAssetShortid: 'template'
        }
      },
      data: {
        name: 'John'
      }
    })

    const text = await textract('test.docx', result.content)
    text.should.containEql('Hello world John')
  })

  it('preview request should return html', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace.docx'))
          }
        }
      },
      data: {
        name: 'John'
      },
      options: {
        preview: true
      }
    })

    result.content.toString().should.containEql('iframe')
  })
})

describe('docx with extensions.docx.previewInWordOnline === false', () => {
  let reporter

  beforeEach(() => {
    reporter = jsreport({
      templatingEngines: {
        strategy: 'in-process'
      }
    }).use(require('../')({ previewInWordOnline: false }))
      .use(require('jsreport-handlebars')())
      .use(require('jsreport-templates')())
      .use(require('jsreport-assets')())
    return reporter.init()
  })

  afterEach(() => reporter.close())

  it('preview request should not return html', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace.docx'))
          }
        }
      },
      data: {
        name: 'John'
      },
      options: {
        preview: true
      }
    })

    result.content.toString().should.not.containEql('iframe')
  })
})
