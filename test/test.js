const should = require('should')
const nock = require('nock')
const jsreport = require('jsreport-core')
const fs = require('fs')
const path = require('path')
const util = require('util')
const { DOMParser } = require('xmldom')
const moment = require('moment')
const toExcelDate = require('js-excel-date-convert').toExcelDate
const { decompress } = require('jsreport-office')
const sizeOf = require('image-size')
const textract = util.promisify(require('textract').fromBufferWithName)
const { nodeListToArray, pxToEMU, cmToEMU } = require('../lib/utils')

async function getImageSize (buf) {
  const files = await decompress()(buf)
  const doc = new DOMParser().parseFromString(
    files.find(f => f.path === 'word/document.xml').data.toString()
  )
  const drawingEl = doc.getElementsByTagName('w:drawing')[0]
  const pictureEl = findDirectPictureChild(drawingEl)
  const aExtEl = pictureEl.getElementsByTagName('a:xfrm')[0].getElementsByTagName('a:ext')[0]

  return {
    width: parseFloat(aExtEl.getAttribute('cx')),
    height: parseFloat(aExtEl.getAttribute('cy'))
  }
}

function findDirectPictureChild (parentNode) {
  const childNodes = parentNode.childNodes || []
  let pictureEl

  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i]

    if (child.nodeName === 'w:drawing') {
      break
    }

    if (child.nodeName === 'pic:pic') {
      pictureEl = child
      break
    }

    const foundInChild = findDirectPictureChild(child)

    if (foundInChild) {
      pictureEl = foundInChild
      break
    }
  }

  return pictureEl
}

describe('docx', () => {
  let reporter

  beforeEach(() => {
    reporter = jsreport({
      store: {
        provider: 'memory'
      },
      templatingEngines: {
        strategy: 'in-process'
      }
    })
      .use(require('../')())
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
            content: fs.readFileSync(
              path.join(__dirname, 'condition-with-helper-call.docx')
            )
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
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace.docx')
            )
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

  it('variable-replace-multi', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace-multi.docx')
            )
          }
        }
      },
      data: {
        name: 'John',
        lastname: 'Wick'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql(
      'Hello world John developer Another lines John developer with Wick as lastname'
    )
  })

  it('variable-replace-syntax-error', () => {
    const prom = reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace-syntax-error.docx')
            )
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    return Promise.all([
      should(prom).be.rejectedWith(/Parse error/),
      // this text that error contains proper location of syntax error
      should(prom).be.rejectedWith(/<w:t>{{<\/w:t>/)
    ])
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
        items: [
          {
            product: {
              name: 'jsreport',
              price: 11
            },
            quantity: 10,
            cost: 20
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('T-123')
    text.should.containEql('jsreport')
    text.should.containEql('Prague 345')
  })

  it('endnote', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'end-note.docx'))
          }
        }
      },
      data: {
        value: 'endnotevalue'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('endnotevalue')
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

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )
    const hyperlink = doc.getElementsByTagName('w:hyperlink')[0]
    const docRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/document.xml.rels').data.toString()
    )
    const rels = nodeListToArray(docRels.getElementsByTagName('Relationship'))

    rels
      .find(node => node.getAttribute('Id') === hyperlink.getAttribute('r:id'))
      .getAttribute('Target')
      .should.be.eql('https://jsreport.net')
  })

  it('link in header', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'link-header.docx'))
          }
        }
      },
      data: {
        linkText: 'jsreport',
        linkUrl: 'https://jsreport.net'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('jsreport')

    const files = await decompress()(result.content)
    const header = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/header1.xml').data.toString()
    )
    const hyperlink = header.getElementsByTagName('w:hyperlink')[0]
    const headerRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/header1.xml.rels').data.toString()
    )
    const rels = nodeListToArray(
      headerRels.getElementsByTagName('Relationship')
    )

    rels
      .find(node => node.getAttribute('Id') === hyperlink.getAttribute('r:id'))
      .getAttribute('Target')
      .should.be.eql('https://jsreport.net')
  })

  it('link in footer', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'link-footer.docx'))
          }
        }
      },
      data: {
        linkText: 'jsreport',
        linkUrl: 'https://jsreport.net'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('jsreport')

    const files = await decompress()(result.content)
    const footer = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/footer1.xml').data.toString()
    )
    const hyperlink = footer.getElementsByTagName('w:hyperlink')[0]
    const footerRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/footer1.xml.rels').data.toString()
    )
    const rels = nodeListToArray(
      footerRels.getElementsByTagName('Relationship')
    )

    rels
      .find(node => node.getAttribute('Id') === hyperlink.getAttribute('r:id'))
      .getAttribute('Target')
      .should.be.eql('https://jsreport.net')
  })

  it('link in header, footer', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'link-header-footer.docx')
            )
          }
        }
      },
      data: {
        linkText: 'jsreport',
        linkUrl: 'https://jsreport.net',
        linkText2: 'github',
        linkUrl2: 'https://github.com'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('jsreport')

    const files = await decompress()(result.content)
    const header = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/header2.xml').data.toString()
    )
    const footer = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/footer2.xml').data.toString()
    )
    const headerHyperlink = header.getElementsByTagName('w:hyperlink')[0]
    const footerHyperlink = footer.getElementsByTagName('w:hyperlink')[0]
    const headerRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/header2.xml.rels').data.toString()
    )
    const footerRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/footer2.xml.rels').data.toString()
    )
    const rels = nodeListToArray(
      headerRels.getElementsByTagName('Relationship')
    )
    const rels2 = nodeListToArray(
      footerRels.getElementsByTagName('Relationship')
    )

    rels
      .find(
        node => node.getAttribute('Id') === headerHyperlink.getAttribute('r:id')
      )
      .getAttribute('Target')
      .should.be.eql('https://jsreport.net')
    rels2
      .find(
        node => node.getAttribute('Id') === footerHyperlink.getAttribute('r:id')
      )
      .getAttribute('Target')
      .should.be.eql('https://github.com')
  })

  it('link to bookmark should not break', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'link-to-bookmark.docx'))
          }
        }
      },
      data: {
        acn: '2222222',
        companyName: 'Demo'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('1 Preliminary 1')
    text.should.containEql('1.1 Name of the Company 1')
    text.should.containEql('1.2 Type of Company 1')
    text.should.containEql('1.3 Limited liability of Members 1')
    text.should.containEql('1.4 The Guarantee 1')
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

    const files = await decompress()(result.content)
    let header1 = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/header1.xml').data.toString()
    )
    let header2 = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/header2.xml').data.toString()
    )
    let header3 = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/header3.xml').data.toString()
    )

    header1
      .getElementsByTagName('v:shape')[0]
      .getElementsByTagName('v:textpath')[0]
      .getAttribute('string')
      .should.be.eql('replacedvalue')
    header2
      .getElementsByTagName('v:shape')[0]
      .getElementsByTagName('v:textpath')[0]
      .getAttribute('string')
      .should.be.eql('replacedvalue')
    header3
      .getElementsByTagName('v:shape')[0]
      .getElementsByTagName('v:textpath')[0]
      .getAttribute('string')
      .should.be.eql('replacedvalue')
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
        people: [
          {
            name: 'Jan'
          },
          {
            name: 'Boris'
          },
          {
            name: 'Pavel'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Jan')
    text.should.containEql('Boris')
  })

  it('list and links', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'list-and-links.docx')
            )
          }
        }
      },
      data: {
        items: [
          {
            text: 'jsreport',
            address: 'https://jsreport.net'
          },
          {
            text: 'github',
            address: 'https://github.com'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('jsreport')
    text.should.containEql('github')
  })

  it('list and endnotes', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'list-and-endnotes.docx')
            )
          }
        }
      },
      data: {
        items: [
          {
            name: '1',
            note: '1n'
          },
          {
            name: '2',
            note: '2n'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('note 1n')
    text.should.containEql('note 2n')
  })

  it('list and footnotes', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'list-and-footnotes.docx')
            )
          }
        }
      },
      data: {
        items: [
          {
            name: '1',
            note: '1n'
          },
          {
            name: '2',
            note: '2n'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('note 1n')
    text.should.containEql('note 2n')
  })

  it('list nested', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'list-nested.docx'))
          }
        }
      },
      data: {
        items: [{
          name: 'Boris',
          items: [{
            name: 'item1'
          }, {
            name: 'item2'
          }]
        }, {
          name: 'Junior',
          items: [{
            name: 'item3'
          }, {
            name: 'item4'
          }]
        }, {
          name: 'Jan',
          items: [{
            name: 'item5'
          }, {
            name: 'item6'
          }]
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const text = await textract('test.docx', result.content)

    text.should.containEql('Boris')
    text.should.containEql('Junior')
    text.should.containEql('Jan')
    text.should.containEql('item1')
    text.should.containEql('item2')
    text.should.containEql('item3')
    text.should.containEql('item4')
    text.should.containEql('item5')
    text.should.containEql('item6')
  })

  it('variable-replace-and-list-after', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace-and-list-after.docx')
            )
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql(
      'This is a test John here we go Test 1 Test 2 Test 3'
    )
  })

  it('variable-replace-and-list-after2', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace-and-list-after2.docx')
            )
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql(
      'This is a test John here we go Test 1 Test 2 Test 3 This is another test John can you see me here'
    )
  })

  it('variable-replace-and-list-after-syntax-error', async () => {
    const prom = reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(
                __dirname,
                'variable-replace-and-list-after-syntax-error.docx'
              )
            )
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    return Promise.all([
      should(prom).be.rejectedWith(/Parse error/),
      // this text that error contains proper location of syntax error
      should(prom).be.rejectedWith(/<w:t>{{<\/w:t>/)
    ])
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
        people: [
          {
            name: 'Jan',
            email: 'jan.blaha@foo.com'
          },
          {
            name: 'Boris',
            email: 'boris@foo.met'
          },
          {
            name: 'Pavel',
            email: 'pavel@foo.met'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Jan')
    text.should.containEql('Boris')
  })

  it('table and links', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'table-and-links.docx')
            )
          }
        }
      },
      data: {
        courses: [
          {
            name: 'The Open University',
            description:
              'Distance and online courses. Qualifications range from certificates, diplomas and short courses to undergraduate and postgraduate degrees.',
            linkName: 'Go to the site1',
            linkURL: 'http://www.openuniversity.edu/courses'
          },
          {
            name: 'Coursera',
            description:
              'Online courses from top universities like Yale, Michigan, Stanford, and leading companies like Google and IBM.',
            linkName: 'Go to the site2',
            linkURL: 'https://plato.stanford.edu/'
          },
          {
            name: 'edX',
            description:
              'Flexible learning on your schedule. Access more than 1900 online courses from 100+ leading institutions including Harvard, MIT, Microsoft, and more.',
            linkName: 'Go to the site3',
            linkURL: 'https://www.edx.org/'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('Go to the site1')
    text.should.containEql('Go to the site2')
    text.should.containEql('Go to the site3')
  })

  it('table and endnotes', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'table-and-endnotes.docx')
            )
          }
        }
      },
      data: {
        courses: [
          {
            name: 'The Open University',
            description:
              'Distance and online courses. Qualifications range from certificates, diplomas and short courses to undergraduate and postgraduate degrees.',
            linkName: 'Go to the site1',
            linkURL: 'http://www.openuniversity.edu/courses',
            note: 'note site1'
          },
          {
            name: 'Coursera',
            description:
              'Online courses from top universities like Yale, Michigan, Stanford, and leading companies like Google and IBM.',
            linkName: 'Go to the site2',
            linkURL: 'https://plato.stanford.edu/',
            note: 'note site2'
          },
          {
            name: 'edX',
            description:
              'Flexible learning on your schedule. Access more than 1900 online courses from 100+ leading institutions including Harvard, MIT, Microsoft, and more.',
            linkName: 'Go to the site3',
            linkURL: 'https://www.edx.org/',
            note: 'note site3'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('note site1')
    text.should.containEql('note site2')
    text.should.containEql('note site3')
  })

  it('table and footnotes', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'table-and-footnotes.docx')
            )
          }
        }
      },
      data: {
        courses: [
          {
            name: 'The Open University',
            description:
              'Distance and online courses. Qualifications range from certificates, diplomas and short courses to undergraduate and postgraduate degrees.',
            linkName: 'Go to the site1',
            linkURL: 'http://www.openuniversity.edu/courses',
            note: 'note site1'
          },
          {
            name: 'Coursera',
            description:
              'Online courses from top universities like Yale, Michigan, Stanford, and leading companies like Google and IBM.',
            linkName: 'Go to the site2',
            linkURL: 'https://plato.stanford.edu/',
            note: 'note site2'
          },
          {
            name: 'edX',
            description:
              'Flexible learning on your schedule. Access more than 1900 online courses from 100+ leading institutions including Harvard, MIT, Microsoft, and more.',
            linkName: 'Go to the site3',
            linkURL: 'https://www.edx.org/',
            note: 'note site3'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('note site1')
    text.should.containEql('note site2')
    text.should.containEql('note site3')
  })

  it('table nested', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'table-nested.docx'))
          }
        }
      },
      data: {
        people: [{
          name: 'Rick',
          lastname: 'Grimes',
          courses: [{
            name: 'Math1',
            homeroom: '2389'
          }, {
            name: 'Math2',
            homeroom: '3389'
          }],
          age: 38
        }, {
          name: 'Andrea',
          lastname: 'Henderson',
          courses: [{
            name: 'Literature1',
            homeroom: '5262'
          }, {
            name: 'Literature2',
            homeroom: '1693'
          }],
          age: 33
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const text = await textract('test.docx', result.content)

    text.should.containEql('Rick')
    text.should.containEql('Andrea')
    text.should.containEql('Math1')
    text.should.containEql('Math2')
    text.should.containEql('Literature1')
    text.should.containEql('Literature2')
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
    const imageBuf = fs.readFileSync(path.join(__dirname, 'image.png'))
    const imageDimensions = sizeOf(imageBuf)

    const targetImageSize = {
      width: pxToEMU(imageDimensions.width),
      height: pxToEMU(imageDimensions.height)
    }

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
        src: 'data:image/png;base64,' + imageBuf.toString('base64')
      }
    })

    const ouputImageSize = await getImageSize(result.content)

    // should preserve original image size by default
    ouputImageSize.width.should.be.eql(targetImageSize.width)
    ouputImageSize.height.should.be.eql(targetImageSize.height)

    fs.writeFileSync('out.docx', result.content)
  })

  it('image with placeholder size (usePlaceholderSize)', async () => {
    const docxBuf = fs.readFileSync(
      path.join(__dirname, 'image-use-placeholder-size.docx')
    )

    let placeholderImageSize

    placeholderImageSize = await getImageSize(docxBuf)

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: docxBuf
          }
        }
      },
      data: {
        src:
          'data:image/png;base64,' +
          fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
      }
    })

    const ouputImageSize = await getImageSize(result.content)

    ouputImageSize.width.should.be.eql(placeholderImageSize.width)
    ouputImageSize.height.should.be.eql(placeholderImageSize.height)

    fs.writeFileSync('out.docx', result.content)
  })

  const units = ['cm', 'px']

  units.forEach(unit => {
    describe(`image size in ${unit}`, () => {
      it('image with custom size (width, height)', async () => {
        const docxBuf = fs.readFileSync(
          path.join(
            __dirname,
            unit === 'cm'
              ? 'image-custom-size.docx'
              : 'image-custom-size-px.docx'
          )
        )

        // 3cm defined in the docx
        const targetImageSize = {
          width: unit === 'cm' ? cmToEMU(3) : pxToEMU(100),
          height: unit === 'cm' ? cmToEMU(3) : pxToEMU(100)
        }

        const result = await reporter.render({
          template: {
            engine: 'handlebars',
            recipe: 'docx',
            docx: {
              templateAsset: {
                content: docxBuf
              }
            }
          },
          data: {
            src:
              'data:image/png;base64,' +
              fs
                .readFileSync(path.join(__dirname, 'image.png'))
                .toString('base64')
          }
        })

        const ouputImageSize = await getImageSize(result.content)

        ouputImageSize.width.should.be.eql(targetImageSize.width)
        ouputImageSize.height.should.be.eql(targetImageSize.height)

        fs.writeFileSync('out.docx', result.content)
      })

      it('image with custom size (width set and height automatic - keep aspect ratio)', async () => {
        const docxBuf = fs.readFileSync(
          path.join(
            __dirname,
            unit === 'cm'
              ? 'image-custom-size-width.docx'
              : 'image-custom-size-width-px.docx'
          )
        )

        const targetImageSize = {
          // 2cm defined in the docx
          width: unit === 'cm' ? cmToEMU(2) : pxToEMU(100),
          // height is calculated automatically based on aspect ratio of image
          height:
            unit === 'cm'
              ? cmToEMU(0.5142851308524194)
              : pxToEMU(25.714330708661418)
        }

        const result = await reporter.render({
          template: {
            engine: 'handlebars',
            recipe: 'docx',
            docx: {
              templateAsset: {
                content: docxBuf
              }
            }
          },
          data: {
            src:
              'data:image/png;base64,' +
              fs
                .readFileSync(path.join(__dirname, 'image.png'))
                .toString('base64')
          }
        })

        const ouputImageSize = await getImageSize(result.content)

        ouputImageSize.width.should.be.eql(targetImageSize.width)
        ouputImageSize.height.should.be.eql(targetImageSize.height)

        fs.writeFileSync('out.docx', result.content)
      })

      it('image with custom size (height set and width automatic - keep aspect ratio)', async () => {
        const docxBuf = fs.readFileSync(
          path.join(
            __dirname,
            unit === 'cm'
              ? 'image-custom-size-height.docx'
              : 'image-custom-size-height-px.docx'
          )
        )

        const targetImageSize = {
          // width is calculated automatically based on aspect ratio of image
          width:
            unit === 'cm'
              ? cmToEMU(7.777781879962101)
              : pxToEMU(194.4444094488189),
          // 2cm defined in the docx
          height: unit === 'cm' ? cmToEMU(2) : pxToEMU(50)
        }

        const result = await reporter.render({
          template: {
            engine: 'handlebars',
            recipe: 'docx',
            docx: {
              templateAsset: {
                content: docxBuf
              }
            }
          },
          data: {
            src:
              'data:image/png;base64,' +
              fs
                .readFileSync(path.join(__dirname, 'image.png'))
                .toString('base64')
          }
        })

        const ouputImageSize = await getImageSize(result.content)

        ouputImageSize.width.should.be.eql(targetImageSize.width)
        ouputImageSize.height.should.be.eql(targetImageSize.height)

        fs.writeFileSync('out.docx', result.content)
      })
    })
  })

  it('image with hyperlink inside', async () => {
    const imageBuf = fs.readFileSync(path.join(__dirname, 'image.png'))

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image-with-hyperlink.docx'))
          }
        }
      },
      data: {
        src: 'data:image/png;base64,' + imageBuf.toString('base64'),
        url: 'https://jsreport.net'
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const drawningEls = doc.getElementsByTagName('w:drawing')

    drawningEls.length.should.be.eql(1)

    const drawningEl = drawningEls[0]

    const isImg = drawningEl.getElementsByTagName('pic:pic').length > 0

    isImg.should.be.True()

    const elLinkClick = drawningEl.getElementsByTagName('a:hlinkClick')[0]
    const hyperlinkRelId = elLinkClick.getAttribute('r:id')

    const docRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/document.xml.rels').data.toString()
    )

    const hyperlinkRelEl = nodeListToArray(docRels.getElementsByTagName('Relationship')).find((el) => {
      return el.getAttribute('Id') === hyperlinkRelId
    })

    const target = decodeURIComponent(hyperlinkRelEl.getAttribute('Target'))

    target.should.be.eql('https://jsreport.net')
  })

  it('image error message when no src provided', async () => {
    return reporter
      .render({
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
      })
      .should.be.rejectedWith(/src parameter to be set/)
  })

  it('image can render from url', async () => {
    const url = 'https://some-server.com/some-image.png'

    nock('https://some-server.com')
      .get('/some-image.png')
      .replyWithFile(200, path.join(__dirname, 'image.png'), {
        'content-type': 'image/png'
      })

    return reporter
      .render({
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
          src: url
        }
      })
      .should.not.be.rejectedWith(/src parameter to be set/)
  })

  it('image error message when src not valid param', async () => {
    return reporter
      .render({
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
      })
      .should.be.rejectedWith(
        /docxImage helper requires src parameter to be valid data uri/
      )
  })

  it('image error message when width not valid param', async () => {
    return reporter
      .render({
        template: {
          engine: 'handlebars',
          recipe: 'docx',
          docx: {
            templateAsset: {
              content: fs.readFileSync(
                path.join(__dirname, 'image-with-wrong-width.docx')
              )
            }
          }
        },
        data: {
          src:
            'data:image/png;base64,' +
            fs
              .readFileSync(path.join(__dirname, 'image.png'))
              .toString('base64')
        }
      })
      .should.be.rejectedWith(
        /docxImage helper requires width parameter to be valid number with unit/
      )
  })

  it('image error message when height not valid param', async () => {
    return reporter
      .render({
        template: {
          engine: 'handlebars',
          recipe: 'docx',
          docx: {
            templateAsset: {
              content: fs.readFileSync(
                path.join(__dirname, 'image-with-wrong-height.docx')
              )
            }
          }
        },
        data: {
          src:
            'data:image/png;base64,' +
            fs
              .readFileSync(path.join(__dirname, 'image.png'))
              .toString('base64')
        }
      })
      .should.be.rejectedWith(
        /docxImage helper requires height parameter to be valid number with unit/
      )
  })

  it('image loop', async () => {
    const images = [
      fs.readFileSync(path.join(__dirname, 'image.png')),
      fs.readFileSync(path.join(__dirname, 'image2.png'))
    ]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image-loop.docx'))
          }
        }
      },
      data: {
        photos: images.map((imageBuf) => {
          return {
            src: 'data:image/png;base64,' + imageBuf.toString('base64')
          }
        })
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const docRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/document.xml.rels').data.toString()
    )

    const drawningEls = nodeListToArray(doc.getElementsByTagName('w:drawing'))

    drawningEls.length.should.be.eql(2)

    drawningEls.forEach((drawningEl, idx) => {
      const isImg = drawningEl.getElementsByTagName('pic:pic').length > 0

      isImg.should.be.True()

      const imageRelId = drawningEl.getElementsByTagName('a:blip')[0].getAttribute('r:embed')

      const imageRelEl = nodeListToArray(docRels.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Id') === imageRelId
      })

      imageRelEl.getAttribute('Type').should.be.eql('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')

      const imageFile = files.find(f => f.path === `word/${imageRelEl.getAttribute('Target')}`)

      // compare returns 0 when buffers are equal
      Buffer.compare(imageFile.data, images[idx]).should.be.eql(0)
    })
  })

  it('image loop and hyperlink inside', async () => {
    const images = [
      {
        url: 'https://jsreport.net',
        buf: fs.readFileSync(path.join(__dirname, 'image.png'))
      },
      {
        url: 'https://www.google.com/intl/es-419/chrome/',
        buf: fs.readFileSync(path.join(__dirname, 'image2.png'))
      }
    ]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image-loop-url.docx'))
          }
        }
      },
      data: {
        photos: images.map((image) => {
          return {
            src: 'data:image/png;base64,' + image.buf.toString('base64'),
            url: image.url
          }
        })
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const docRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/document.xml.rels').data.toString()
    )

    const drawningEls = nodeListToArray(doc.getElementsByTagName('w:drawing'))

    drawningEls.length.should.be.eql(2)

    drawningEls.forEach((drawningEl, idx) => {
      const isImg = drawningEl.getElementsByTagName('pic:pic').length > 0

      isImg.should.be.True()

      const imageRelId = drawningEl.getElementsByTagName('a:blip')[0].getAttribute('r:embed')

      const imageRelEl = nodeListToArray(docRels.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Id') === imageRelId
      })

      imageRelEl.getAttribute('Type').should.be.eql('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')

      const imageFile = files.find(f => f.path === `word/${imageRelEl.getAttribute('Target')}`)

      // compare returns 0 when buffers are equal
      Buffer.compare(imageFile.data, images[idx].buf).should.be.eql(0)

      const elLinkClick = drawningEl.getElementsByTagName('a:hlinkClick')[0]
      const hyperlinkRelId = elLinkClick.getAttribute('r:id')

      const hyperlinkRelEl = nodeListToArray(docRels.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Id') === hyperlinkRelId
      })

      const target = decodeURIComponent(hyperlinkRelEl.getAttribute('Target'))

      target.should.be.eql(images[idx].url)
    })
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
        chapters: [
          {
            title: 'Chapter 1',
            text: 'This is the first chapter'
          },
          {
            title: 'Chapter 2',
            text: 'This is the second chapter'
          }
        ]
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
        experiences: [
          {
            title: '.NET Developer',
            company: 'Unicorn',
            from: '1.1.2010',
            to: '15.5.2012'
          },
          {
            title: 'Solution Architect',
            company: 'Simplias',
            from: '15.5.2012',
            to: 'now'
          }
        ],
        skills: [
          {
            title: 'The worst developer ever'
          },
          {
            title: `Don't need to write semicolons`
          }
        ],
        printFooter: true
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('Jan Blaha')
  })

  it('input form control', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-input.docx')
            )
          }
        }
      },
      data: {
        name: 'Erick'
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc
      .getElementsByTagName('w:textInput')[0]
      .getElementsByTagName('w:default')[0]
      .getAttribute('w:val')
      .should.be.eql('Erick')
  })

  it('checkbox form control', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-checkbox.docx')
            )
          }
        }
      },
      data: {
        ready: true
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc
      .getElementsByTagName('w14:checked')[0]
      .getAttribute('w14:val')
      .should.be.eql('1')

    doc
      .getElementsByTagName('w:sdt')[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('☒')
  })

  it('combobox form control', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-combo.docx')
            )
          }
        }
      },
      data: {
        val: 'vala'
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc.getElementsByTagName('w:sdtContent')[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('display val')
  })

  it('combobox form control with constant value', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-combo-constant-value.docx')
            )
          }
        }
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc.getElementsByTagName('w:sdtContent')[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('value a')
  })

  it('combobox form control with dynamic items', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-combo-dynamic-items.docx')
            )
          }
        }
      },
      data: {
        val: 'b',
        items: [
          {
            value: 'a',
            text: 'Jan'
          },
          {
            value: 'b',
            text: 'Boris'
          }
        ]
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc.getElementsByTagName('w:sdtContent')[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('Boris')
  })

  it('combobox form control with dynamic items in strings', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-combo-dynamic-items.docx')
            )
          }
        }
      },
      data: {
        val: 'Boris',
        items: ['Jan', 'Boris']
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc.getElementsByTagName('w:sdtContent')[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('Boris')
  })

  it('combobox form control with dynamic items in strings and special character', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'form-control-combo-dynamic-items.docx')
            )
          }
        }
      },
      data: {
        val: 'Boris$',
        items: ['Jan$', 'Boris$']
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    doc.getElementsByTagName('w:sdtContent')[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('Boris$')
  })

  it('chart', async () => {
    const labels = ['Jan', 'Feb', 'March']
    const datasets = [{
      label: 'Ser1',
      data: [4, 5, 1]
    }, {
      label: 'Ser2',
      data: [2, 3, 5]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const dataElements = nodeListToArray(doc.getElementsByTagName('c:ser'))

    dataElements.forEach((dataEl, idx) => {
      dataEl.getElementsByTagName('c:tx')[0].getElementsByTagName('c:v')[0].textContent.should.be.eql(datasets[idx].label)
      nodeListToArray(dataEl.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')).map((el) => el.textContent).should.be.eql(labels)
      nodeListToArray(dataEl.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')).map((el) => parseInt(el.textContent, 10)).should.be.eql(datasets[idx].data)
    })
  })

  it('chart without style, color xml files', async () => {
    const labels = ['Q1', 'Q2', 'Q3', 'Q4']
    const datasets = [{
      label: 'Apples',
      data: [100, 50, 10, 70]
    }, {
      label: 'Oranges',
      data: [20, 30, 20, 40]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart-with-no-style-colors-xml-files.docx'))
          }
        }
      },
      data: {
        fruits: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const dataElements = nodeListToArray(doc.getElementsByTagName('c:ser'))

    dataElements.forEach((dataEl, idx) => {
      dataEl.getElementsByTagName('c:tx')[0].getElementsByTagName('c:v')[0].textContent.should.be.eql(datasets[idx].label)
      nodeListToArray(dataEl.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')).map((el) => el.textContent).should.be.eql(labels)
      nodeListToArray(dataEl.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')).map((el) => parseInt(el.textContent, 10)).should.be.eql(datasets[idx].data)
    })
  })

  it('chart with title', async () => {
    const labels = ['Jan', 'Feb', 'March']
    const datasets = [{
      label: 'Ser1',
      data: [4, 5, 1]
    }, {
      label: 'Ser2',
      data: [2, 3, 5]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart-with-title.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const chartTitleEl = doc.getElementsByTagName('c:title')[0].getElementsByTagName('a:t')[0]

    chartTitleEl.textContent.should.be.eql('DEMO CHART')
  })

  it('chart with dynamic title', async () => {
    const labels = ['Jan', 'Feb', 'March']
    const datasets = [{
      label: 'Ser1',
      data: [4, 5, 1]
    }, {
      label: 'Ser2',
      data: [2, 3, 5]
    }]
    const chartTitle = 'CUSTOM CHART'

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart-with-dynamic-title.docx'))
          }
        }
      },
      data: {
        chartTitle,
        chartData: {
          labels,
          datasets
        }
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const chartTitleEl = doc.getElementsByTagName('c:title')[0].getElementsByTagName('a:t')[0]

    chartTitleEl.textContent.should.be.eql(chartTitle)
  })

  it('stock chart', async () => {
    const labels = [
      '2020-05-10',
      '2020-06-10',
      '2020-07-10',
      '2020-08-10'
    ]

    const datasets = [{
      label: 'High',
      data: [
        43,
        56,
        24,
        36
      ]
    }, {
      label: 'Low',
      data: [
        17,
        25,
        47,
        32
      ]
    }, {
      label: 'Close',
      data: [
        19,
        42,
        29,
        33
      ]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'stock-chart.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const dataElements = nodeListToArray(doc.getElementsByTagName('c:ser'))

    dataElements.forEach((dataEl, idx) => {
      dataEl.getElementsByTagName('c:tx')[0].getElementsByTagName('c:v')[0].textContent.should.be.eql(datasets[idx].label)

      nodeListToArray(dataEl.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')).map((el) => el.textContent).should.be.eql(labels.map((l) => {
        return toExcelDate(moment(l).toDate()).toString()
      }))

      nodeListToArray(dataEl.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')).map((el) => parseInt(el.textContent, 10)).should.be.eql(datasets[idx].data)
    })
  })

  it('waterfall chart (chartex)', async () => {
    const labels = [
      'Cat 1',
      'Cat 2',
      'Cat 3',
      'Cat 4',
      'Cat 5',
      'Cat 5',
      'Cat 6',
      'Cat 8',
      'Cat 9'
    ]

    const datasets = [{
      label: 'Water Fall',
      data: [
        9702.0,
        -210.3,
        -24.0,
        -674.0,
        19.4,
        -1406.9,
        352.9,
        2707.5,
        10466.5
      ]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'waterfall-chart.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chartEx1.xml').data.toString()
    )

    const labelElement = (
      doc.getElementsByTagName('cx:series')[0]
        .getElementsByTagName('cx:txData')[0]
        .getElementsByTagName('cx:v')[0]
    )

    const dataElement = doc.getElementsByTagName('cx:data')[0]

    labelElement.textContent.should.be.eql(datasets[0].label)

    const strDimElement = dataElement.getElementsByTagName('cx:strDim')[0]
    const numDimElement = dataElement.getElementsByTagName('cx:numDim')[0]

    strDimElement.getAttribute('type').should.be.eql('cat')
    numDimElement.getAttribute('type').should.be.eql('val')

    nodeListToArray(
      strDimElement
        .getElementsByTagName('cx:lvl')[0]
        .getElementsByTagName('cx:pt')
    ).forEach((dataEl, idx) => {
      dataEl.textContent.should.be.eql(labels[idx])
    })

    nodeListToArray(
      numDimElement
        .getElementsByTagName('cx:lvl')[0]
        .getElementsByTagName('cx:pt')
    ).forEach((dataEl, idx) => {
      parseFloat(dataEl.textContent).should.be.eql(datasets[0].data[idx])
    })
  })

  it('funnel chart (chartex)', async () => {
    const labels = [
      'Cat 1',
      'Cat 2',
      'Cat 3',
      'Cat 4',
      'Cat 5',
      'Cat 6'
    ]

    const datasets = [{
      label: 'Funnel',
      data: [
        3247,
        5729,
        1395,
        2874,
        6582,
        1765
      ]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'funnel-chart.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chartEx1.xml').data.toString()
    )

    const labelElement = (
      doc.getElementsByTagName('cx:series')[0]
        .getElementsByTagName('cx:txData')[0]
        .getElementsByTagName('cx:v')[0]
    )

    const dataElement = doc.getElementsByTagName('cx:data')[0]

    labelElement.textContent.should.be.eql(datasets[0].label)

    const strDimElement = dataElement.getElementsByTagName('cx:strDim')[0]
    const numDimElement = dataElement.getElementsByTagName('cx:numDim')[0]

    strDimElement.getAttribute('type').should.be.eql('cat')
    numDimElement.getAttribute('type').should.be.eql('val')

    nodeListToArray(
      strDimElement
        .getElementsByTagName('cx:lvl')[0]
        .getElementsByTagName('cx:pt')
    ).forEach((dataEl, idx) => {
      dataEl.textContent.should.be.eql(labels[idx])
    })

    nodeListToArray(
      numDimElement
        .getElementsByTagName('cx:lvl')[0]
        .getElementsByTagName('cx:pt')
    ).forEach((dataEl, idx) => {
      parseFloat(dataEl.textContent).should.be.eql(datasets[0].data[idx])
    })
  })

  it('treemap chart (chartex)', async () => {
    const labels = [
      [
        'Rama 1',
        'Rama 1',
        'Rama 1',
        'Rama 1',
        'Rama 1',
        'Rama 2',
        'Rama 2',
        'Rama 3'
      ],
      [
        'Tallo 1',
        'Tallo 1',
        'Tallo 1',
        'Tallo 2',
        'Tallo 2',
        'Tallo 2',
        'Tallo 3',
        'Tallo 3'
      ],
      [
        'Hoja 1',
        'Hoja 2',
        'Hoja 3',
        'Hoja 4',
        'Hoja 5',
        'Hoja 6',
        'Hoja 7',
        'Hoja 8'
      ]
    ]

    const datasets = [{
      label: 'Treemap',
      data: [
        52,
        43,
        56,
        76,
        91,
        49,
        31,
        81
      ]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'treemap-chart.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chartEx1.xml').data.toString()
    )

    const labelElement = (
      doc.getElementsByTagName('cx:series')[0]
        .getElementsByTagName('cx:txData')[0]
        .getElementsByTagName('cx:v')[0]
    )

    const dataElement = doc.getElementsByTagName('cx:data')[0]

    labelElement.textContent.should.be.eql(datasets[0].label)

    const strDimElement = dataElement.getElementsByTagName('cx:strDim')[0]
    const numDimElement = dataElement.getElementsByTagName('cx:numDim')[0]

    strDimElement.getAttribute('type').should.be.eql('cat')
    numDimElement.getAttribute('type').should.be.eql('size')

    nodeListToArray(
      strDimElement.getElementsByTagName('cx:lvl')
    ).forEach((lvlEl, idx) => {
      const targetLabels = labels.reverse()[idx]

      nodeListToArray(lvlEl.getElementsByTagName('cx:pt')).forEach((dataEl, ydx) => {
        dataEl.textContent.should.be.eql(targetLabels[ydx])
      })
    })

    nodeListToArray(
      numDimElement
        .getElementsByTagName('cx:lvl')[0]
        .getElementsByTagName('cx:pt')
    ).forEach((dataEl, idx) => {
      parseFloat(dataEl.textContent).should.be.eql(datasets[0].data[idx])
    })
  })

  it('sunburst chart (chartex)', async () => {
    const labels = [
      [
        'Rama 1',
        'Rama 1',
        'Rama 1',
        'Rama 2',
        'Rama 2',
        'Rama 2',
        'Rama 2',
        'Rama 3'
      ],
      [
        'Tallo 1',
        'Tallo 1',
        'Tallo 1',
        'Tallo 2',
        'Tallo 2',
        'Tallo 2',
        'Hoja 6',
        'Hoja 7'
      ],
      [
        'Hoja 1',
        'Hoja 2',
        'Hoja 3',
        'Hoja 4',
        'Hoja 5',
        null,
        null,
        'Hoja 8'
      ]
    ]

    const datasets = [{
      label: 'Sunburst',
      data: [
        32,
        68,
        83,
        72,
        75,
        84,
        52,
        34
      ]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'sunburst-chart.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chartEx1.xml').data.toString()
    )

    const labelElement = (
      doc.getElementsByTagName('cx:series')[0]
        .getElementsByTagName('cx:txData')[0]
        .getElementsByTagName('cx:v')[0]
    )

    const dataElement = doc.getElementsByTagName('cx:data')[0]

    labelElement.textContent.should.be.eql(datasets[0].label)

    const strDimElement = dataElement.getElementsByTagName('cx:strDim')[0]
    const numDimElement = dataElement.getElementsByTagName('cx:numDim')[0]

    strDimElement.getAttribute('type').should.be.eql('cat')
    numDimElement.getAttribute('type').should.be.eql('size')

    nodeListToArray(
      strDimElement.getElementsByTagName('cx:lvl')
    ).forEach((lvlEl, idx) => {
      const targetLabels = labels.reverse()[idx]

      nodeListToArray(lvlEl.getElementsByTagName('cx:pt')).forEach((dataEl, ydx) => {
        dataEl.textContent.should.be.eql(targetLabels[ydx] || '')
      })
    })

    nodeListToArray(
      numDimElement
        .getElementsByTagName('cx:lvl')[0]
        .getElementsByTagName('cx:pt')
    ).forEach((dataEl, idx) => {
      parseFloat(dataEl.textContent).should.be.eql(datasets[0].data[idx])
    })
  })

  it('chart error message when no data', async () => {
    return reporter
      .render({
        template: {
          engine: 'handlebars',
          recipe: 'docx',
          docx: {
            templateAsset: {
              content: fs.readFileSync(path.join(__dirname, 'chart-error-data.docx'))
            }
          }
        },
        data: {
          chartData: null
        }
      })
      .should.be.rejectedWith(/requires data parameter to be set/)
  })

  it('chart error message when no data.labels', async () => {
    return reporter
      .render({
        template: {
          engine: 'handlebars',
          recipe: 'docx',
          docx: {
            templateAsset: {
              content: fs.readFileSync(path.join(__dirname, 'chart-error-data.docx'))
            }
          }
        },
        data: {
          chartData: {}
        }
      })
      .should.be.rejectedWith(/requires data parameter with labels to be set/)
  })

  it('chart error message when no data.datasets', async () => {
    return reporter
      .render({
        template: {
          engine: 'handlebars',
          recipe: 'docx',
          docx: {
            templateAsset: {
              content: fs.readFileSync(path.join(__dirname, 'chart-error-data.docx'))
            }
          }
        },
        data: {
          chartData: {
            labels: ['Jan', 'Feb', 'March'],
            datasets: null
          }
        }
      })
      .should.be.rejectedWith(/requires data parameter with datasets to be set/)
  })

  it('chart loop', async () => {
    const charts = [{
      chartData: {
        labels: ['Jan', 'Feb', 'March'],
        datasets: [{
          label: 'Ser1',
          data: [4, 5, 1]
        }, {
          label: 'Ser2',
          data: [2, 3, 5]
        }]
      }
    }, {
      chartData: {
        labels: ['Apr', 'May', 'Jun'],
        datasets: [{
          label: 'Ser3',
          data: [8, 2, 4]
        }, {
          label: 'Ser4',
          data: [2, 5, 3]
        }]
      }
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart-loop.docx'))
          }
        }
      },
      data: {
        charts
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const chartDrawningEls = nodeListToArray(doc.getElementsByTagName('c:chart'))

    chartDrawningEls.length.should.be.eql(charts.length)

    const docRels = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/_rels/document.xml.rels').data.toString()
    )

    chartDrawningEls.forEach((chartDrawningEl, chartIdx) => {
      const chartRelId = chartDrawningEl.getAttribute('r:id')

      const chartRelEl = nodeListToArray(docRels.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Id') === chartRelId
      })

      const chartDoc = new DOMParser().parseFromString(
        files.find(f => f.path === `word/${chartRelEl.getAttribute('Target')}`).data.toString()
      )

      const chartRelsDoc = new DOMParser().parseFromString(
        files.find(f => f.path === `word/charts/_rels/${chartRelEl.getAttribute('Target').split('/').slice(-1)[0]}.rels`).data.toString()
      )

      const dataElements = nodeListToArray(chartDoc.getElementsByTagName('c:ser'))

      dataElements.forEach((dataEl, idx) => {
        dataEl.getElementsByTagName('c:tx')[0].getElementsByTagName('c:v')[0].textContent.should.be.eql(charts[chartIdx].chartData.datasets[idx].label)
        nodeListToArray(dataEl.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')).map((el) => el.textContent).should.be.eql(charts[chartIdx].chartData.labels)
        nodeListToArray(dataEl.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')).map((el) => parseInt(el.textContent, 10)).should.be.eql(charts[chartIdx].chartData.datasets[idx].data)
      })

      const chartStyleRelEl = nodeListToArray(chartRelsDoc.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Type') === 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
      })

      const chartStyleDoc = files.find(f => f.path === `word/charts/${chartStyleRelEl.getAttribute('Target')}`)

      chartStyleDoc.should.be.not.undefined()

      const chartColorStyleRelEl = nodeListToArray(chartRelsDoc.getElementsByTagName('Relationship')).find((el) => {
        return el.getAttribute('Type') === 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
      })

      const chartColorStyleDoc = files.find(f => f.path === `word/charts/${chartColorStyleRelEl.getAttribute('Target')}`)

      chartColorStyleDoc.should.be.not.undefined()
    })
  })

  it('chart should keep style defined in serie', async () => {
    const labels = ['Q1', 'Q2', 'Q3', 'Q4']
    const datasets = [{
      label: 'Apples',
      data: [100, 50, 10, 70]
    }, {
      label: 'Oranges',
      data: [20, 30, 20, 40]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart-serie-style.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const dataElements = nodeListToArray(doc.getElementsByTagName('c:ser'))

    dataElements.forEach((dataEl, idx) => {
      should(dataEl.getElementsByTagName('c:spPr')[0]).be.not.undefined()
    })
  })

  it('chart should keep number format defined in serie', async () => {
    const labels = ['Q1', 'Q2', 'Q3', 'Q4']
    const datasets = [{
      label: 'Apples',
      data: [10000.0, 50000.45, 10000.45, 70000.546]
    }, {
      label: 'Oranges',
      data: [20000.3, 30000.2, 20000.4, 40000.4]
    }]

    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'chart-serie-number-format.docx'))
          }
        }
      },
      data: {
        chartData: {
          labels,
          datasets
        }
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/charts/chart1.xml').data.toString()
    )

    const dataElements = nodeListToArray(doc.getElementsByTagName('c:ser'))

    dataElements.forEach((dataEl, idx) => {
      should(dataEl.getElementsByTagName('c:val')[0].getElementsByTagName('c:formatCode')[0].textContent).be.eql('#,##0.0')
    })
  })

  it('should not duplicate drawing object id in loop', async () => {
    // drawing object should not contain duplicated id, otherwhise it produce a warning in ms word
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'dw-object-loop-id.docx'))
          }
        }
      },
      data: {
        items: [1, 2, 3]
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const drawingEls = nodeListToArray(doc.getElementsByTagName('w:drawing'))
    const baseId = 12

    drawingEls.forEach((drawingEl, idx) => {
      const docPrEl = nodeListToArray(drawingEl.firstChild.childNodes).find((el) => el.nodeName === 'wp:docPr')
      parseInt(docPrEl.getAttribute('id'), 10).should.be.eql(baseId + idx)
    })
  })

  it('page break in single paragraph', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'page-break-single-paragraph.docx')
            )
          }
        }
      },
      data: {}
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const paragraphNodes = nodeListToArray(doc.getElementsByTagName('w:p'))

    paragraphNodes[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('Demo')

    paragraphNodes[0].getElementsByTagName('w:br').should.have.length(1)
    paragraphNodes[0]
      .getElementsByTagName('w:t')[1]
      .textContent.should.be.eql('break')
  })

  it('page break in single paragraph (sharing text nodes)', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'page-break-single-paragraph2.docx')
            )
          }
        }
      },
      data: {}
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)

    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const paragraphNodes = nodeListToArray(doc.getElementsByTagName('w:p'))

    paragraphNodes[0]
      .getElementsByTagName('w:t')[0]
      .textContent.should.be.eql('Demo')

    paragraphNodes[0].getElementsByTagName('w:br').should.have.length(1)

    paragraphNodes[0]
      .getElementsByTagName('w:t')[1]
      .textContent.should.be.eql('of a break')
  })

  it('page break between paragraphs', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(
              path.join(__dirname, 'page-break-between-paragraphs.docx')
            )
          }
        }
      },
      data: {}
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(
      files.find(f => f.path === 'word/document.xml').data.toString()
    )

    const paragraphNodes = nodeListToArray(
      doc.getElementsByTagName('w:p')
    ).filter(p => {
      const breakNodes = getBreaks(p)

      const hasText = getText(p) != null && getText(p) !== ''

      if (!hasText && breakNodes.length === 0) {
        return false
      }

      return true
    })

    function getText (p) {
      const textNodes = nodeListToArray(p.getElementsByTagName('w:t')).filter(
        t => {
          return t.textContent != null && t.textContent !== ''
        }
      )

      return textNodes.map(t => t.textContent).join('')
    }

    function getBreaks (p) {
      return nodeListToArray(p.getElementsByTagName('w:br'))
    }

    getText(paragraphNodes[0]).should.be.eql('Demo some text')
    getBreaks(paragraphNodes[1]).should.have.length(1)
    getText(paragraphNodes[2]).should.be.eql('after break')
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
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace.docx')
            )
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
    })
      .use(require('../')({ preview: { enabled: false } }))
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
            content: fs.readFileSync(
              path.join(__dirname, 'variable-replace.docx')
            )
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
