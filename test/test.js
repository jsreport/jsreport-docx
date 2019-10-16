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

  it('condition-with-helper-call', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'condition-with-helper-call.docx'))
          }
        },
        helpers: `
          function moreThan2(users) {
            return users.length > 2
          }
        `
      },
      data: {
        users: [1, 2, 3]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('More than 2 users')
  })

  it('condition with docProps/thumbnail.jpeg in docx', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'condition.docx'))
          }
        },
        helpers: `
          function moreThan2(users) {
            return users.length > 2
          }
        `
      },
      data: {
        users: [1, 2, 3]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('More than 2 users')
  })

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

  it('invoice', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'invoice.docx'))
          }
        }
      },
      data: {
        invoiceNumber: 'T-123',
        company: {
          address: 'Prague 345',
          email: 'foo',
          phone: 'phone'
        },
        total: 1000,
        date: 'dddd',
        items: [{
          product: {
            name: 'jsreport',
            price: 11
          },
          quantity: 10,
          cost: 20
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('T-123')
    text.should.containEql('jsreport')
    text.should.containEql('Prague 345')
  })

  it('footnote', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'foot-note.docx'))
          }
        }
      },
      data: {
        value: 'footnotevalue'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('footnotevalue')
  })

  it('link', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'link.docx'))
          }
        }
      },
      data: {
        url: 'https://jsreport.net'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('website')
  })

  it('watermark', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'watermark.docx'))
          }
        }
      },
      data: {
        watermark: 'replacedvalue'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    // text parser is not parsing watermarks
  })

  it('list', async () => {
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

  it('table', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'table.docx'))
          }
        }
      },
      data: {
        people: [{
          name: 'Jan', email: 'jan.blaha@foo.com'
        }, {
          name: 'Boris', email: 'boris@foo.met'
        }, {
          name: 'Pavel', email: 'pavel@foo.met'
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Jan')
    text.should.containEql('Boris')
  })

  it('style', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'style.docx'))
          }
        }
      },
      data: {}
    })

    fs.writeFileSync('out.docx', result.content)
  })

  it('image', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image.docx'))
          }
        }
      },
      data: {
        src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
      }
    })

    fs.writeFileSync('out.docx', result.content)
  })

  it('image error message when no src provided', async () => {
    return reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image.docx'))
          }
        }
      },
      data: {
        src: null
      }
    }).should.be.rejectedWith(/src parameter to be set/)
  })

  it('image error message when src not valid param', async () => {
    return reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image.docx'))
          }
        }
      },
      data: {
        src: 'data:image/gif;base64,R0lG'
      }
    }).should.be.rejectedWith(/docxImage helper requires src parameter to be valid data uri/)
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
        },
        helpers: `
          function customHelper(options) {
            return options.fn(this)
          }
        `
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
        }],
        printFooter: true
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
    }).use(require('../')({ preview: { enabled: false } }))
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
