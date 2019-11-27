const should = require('should')
const jsreport = require('jsreport-core')
const fs = require('fs')
const path = require('path')
const util = require('util')
const { DOMParser } = require('xmldom')
const { decompress } = require('jsreport-office')
const sizeOf = require('image-size')
const textract = util.promisify(require('textract').fromBufferWithName)
const { nodeListToArray, pxToEMU, cmToEMU } = require('../lib/utils')

async function getImageSize (buf) {
  const files = await decompress()(buf)
  const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())
  const elDrawing = doc.getElementsByTagName('w:drawing')[0]
  const wpExtendEl = elDrawing.getElementsByTagName('wp:extent')[0]

  return {
    width: parseFloat(wpExtendEl.getAttribute('cx')),
    height: parseFloat(wpExtendEl.getAttribute('cy'))
  }
}

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

  it('variable-replace-multi', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace-multi.docx'))
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
    text.should.containEql('Hello world John developer Another lines John developer with Wick as lastname')
  })

  it('variable-replace-syntax-error', () => {
    const prom = reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace-syntax-error.docx'))
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
    const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())
    const hyperlink = doc.getElementsByTagName('w:hyperlink')[0]
    const docRels = new DOMParser().parseFromString(files.find(f => f.path === 'word/_rels/document.xml.rels').data.toString())
    const rels = nodeListToArray(docRels.getElementsByTagName('Relationship'))

    rels.find((node) => node.getAttribute('Id') === hyperlink.getAttribute('r:id')).getAttribute('Target').should.be.eql('https://jsreport.net')
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
    const header = new DOMParser().parseFromString(files.find(f => f.path === 'word/header1.xml').data.toString())
    const hyperlink = header.getElementsByTagName('w:hyperlink')[0]
    const headerRels = new DOMParser().parseFromString(files.find(f => f.path === 'word/_rels/header1.xml.rels').data.toString())
    const rels = nodeListToArray(headerRels.getElementsByTagName('Relationship'))

    rels.find((node) => node.getAttribute('Id') === hyperlink.getAttribute('r:id')).getAttribute('Target').should.be.eql('https://jsreport.net')
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
    const footer = new DOMParser().parseFromString(files.find(f => f.path === 'word/footer1.xml').data.toString())
    const hyperlink = footer.getElementsByTagName('w:hyperlink')[0]
    const footerRels = new DOMParser().parseFromString(files.find(f => f.path === 'word/_rels/footer1.xml.rels').data.toString())
    const rels = nodeListToArray(footerRels.getElementsByTagName('Relationship'))

    rels.find((node) => node.getAttribute('Id') === hyperlink.getAttribute('r:id')).getAttribute('Target').should.be.eql('https://jsreport.net')
  })

  it('link in header, footer', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'link-header-footer.docx'))
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
    const header = new DOMParser().parseFromString(files.find(f => f.path === 'word/header2.xml').data.toString())
    const footer = new DOMParser().parseFromString(files.find(f => f.path === 'word/footer2.xml').data.toString())
    const headerHyperlink = header.getElementsByTagName('w:hyperlink')[0]
    const footerHyperlink = footer.getElementsByTagName('w:hyperlink')[0]
    const headerRels = new DOMParser().parseFromString(files.find(f => f.path === 'word/_rels/header2.xml.rels').data.toString())
    const footerRels = new DOMParser().parseFromString(files.find(f => f.path === 'word/_rels/footer2.xml.rels').data.toString())
    const rels = nodeListToArray(headerRels.getElementsByTagName('Relationship'))
    const rels2 = nodeListToArray(footerRels.getElementsByTagName('Relationship'))

    rels.find((node) => node.getAttribute('Id') === headerHyperlink.getAttribute('r:id')).getAttribute('Target').should.be.eql('https://jsreport.net')
    rels2.find((node) => node.getAttribute('Id') === footerHyperlink.getAttribute('r:id')).getAttribute('Target').should.be.eql('https://github.com')
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
    let header1 = new DOMParser().parseFromString(files.find(f => f.path === 'word/header1.xml').data.toString())
    let header2 = new DOMParser().parseFromString(files.find(f => f.path === 'word/header2.xml').data.toString())
    let header3 = new DOMParser().parseFromString(files.find(f => f.path === 'word/header3.xml').data.toString())

    header1.getElementsByTagName('v:shape')[0].getElementsByTagName('v:textpath')[0].getAttribute('string').should.be.eql('replacedvalue')
    header2.getElementsByTagName('v:shape')[0].getElementsByTagName('v:textpath')[0].getAttribute('string').should.be.eql('replacedvalue')
    header3.getElementsByTagName('v:shape')[0].getElementsByTagName('v:textpath')[0].getAttribute('string').should.be.eql('replacedvalue')
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

  it('list and links', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'list-and-links.docx'))
          }
        }
      },
      data: {
        items: [{
          text: 'jsreport',
          address: 'https://jsreport.net'
        }, {
          text: 'github',
          address: 'https://github.com'
        }]
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
            content: fs.readFileSync(path.join(__dirname, 'list-and-endnotes.docx'))
          }
        }
      },
      data: {
        items: [{
          name: '1',
          note: '1n'
        }, {
          name: '2',
          note: '2n'
        }]
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
            content: fs.readFileSync(path.join(__dirname, 'list-and-footnotes.docx'))
          }
        }
      },
      data: {
        items: [{
          name: '1',
          note: '1n'
        }, {
          name: '2',
          note: '2n'
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('note 1n')
    text.should.containEql('note 2n')
  })

  it('variable-replace-and-list-after', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace-and-list-after.docx'))
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('This is a test John here we go Test 1 Test 2 Test 3')
  })

  it('variable-replace-and-list-after2', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace-and-list-after2.docx'))
          }
        }
      },
      data: {
        name: 'John'
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)
    text.should.containEql('This is a test John here we go Test 1 Test 2 Test 3 This is another test John can you see me here')
  })

  it('variable-replace-and-list-after-syntax-error', async () => {
    const prom = reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'variable-replace-and-list-after-syntax-error.docx'))
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

  it('table and links', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'table-and-links.docx'))
          }
        }
      },
      data: {
        courses: [{
          name: 'The Open University',
          description: 'Distance and online courses. Qualifications range from certificates, diplomas and short courses to undergraduate and postgraduate degrees.',
          linkName: 'Go to the site1',
          linkURL: 'http://www.openuniversity.edu/courses'
        }, {
          name: 'Coursera',
          description: 'Online courses from top universities like Yale, Michigan, Stanford, and leading companies like Google and IBM.',
          linkName: 'Go to the site2',
          linkURL: 'https://plato.stanford.edu/'
        }, {
          name: 'edX',
          description: 'Flexible learning on your schedule. Access more than 1900 online courses from 100+ leading institutions including Harvard, MIT, Microsoft, and more.',
          linkName: 'Go to the site3',
          linkURL: 'https://www.edx.org/'
        }]
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
            content: fs.readFileSync(path.join(__dirname, 'table-and-endnotes.docx'))
          }
        }
      },
      data: {
        courses: [{
          name: 'The Open University',
          description: 'Distance and online courses. Qualifications range from certificates, diplomas and short courses to undergraduate and postgraduate degrees.',
          linkName: 'Go to the site1',
          linkURL: 'http://www.openuniversity.edu/courses',
          note: 'note site1'
        }, {
          name: 'Coursera',
          description: 'Online courses from top universities like Yale, Michigan, Stanford, and leading companies like Google and IBM.',
          linkName: 'Go to the site2',
          linkURL: 'https://plato.stanford.edu/',
          note: 'note site2'
        }, {
          name: 'edX',
          description: 'Flexible learning on your schedule. Access more than 1900 online courses from 100+ leading institutions including Harvard, MIT, Microsoft, and more.',
          linkName: 'Go to the site3',
          linkURL: 'https://www.edx.org/',
          note: 'note site3'
        }]
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
            content: fs.readFileSync(path.join(__dirname, 'table-and-footnotes.docx'))
          }
        }
      },
      data: {
        courses: [{
          name: 'The Open University',
          description: 'Distance and online courses. Qualifications range from certificates, diplomas and short courses to undergraduate and postgraduate degrees.',
          linkName: 'Go to the site1',
          linkURL: 'http://www.openuniversity.edu/courses',
          note: 'note site1'
        }, {
          name: 'Coursera',
          description: 'Online courses from top universities like Yale, Michigan, Stanford, and leading companies like Google and IBM.',
          linkName: 'Go to the site2',
          linkURL: 'https://plato.stanford.edu/',
          note: 'note site2'
        }, {
          name: 'edX',
          description: 'Flexible learning on your schedule. Access more than 1900 online courses from 100+ leading institutions including Harvard, MIT, Microsoft, and more.',
          linkName: 'Go to the site3',
          linkURL: 'https://www.edx.org/',
          note: 'note site3'
        }]
      }
    })

    fs.writeFileSync('out.docx', result.content)
    const text = await textract('test.docx', result.content)

    text.should.containEql('note site1')
    text.should.containEql('note site2')
    text.should.containEql('note site3')
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
    const docxBuf = fs.readFileSync(path.join(__dirname, 'image-use-placeholder-size.docx'))

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
        src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
      }
    })

    const ouputImageSize = await getImageSize(result.content)

    ouputImageSize.width.should.be.eql(placeholderImageSize.width)
    ouputImageSize.height.should.be.eql(placeholderImageSize.height)

    fs.writeFileSync('out.docx', result.content)
  })

  const units = ['cm', 'px']

  units.forEach((unit) => {
    describe(`image size in ${unit}`, () => {
      it('image with custom size (width, height)', async () => {
        const docxBuf = fs.readFileSync(path.join(__dirname, unit === 'cm' ? 'image-custom-size.docx' : 'image-custom-size-px.docx'))

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
            src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
          }
        })

        const ouputImageSize = await getImageSize(result.content)

        ouputImageSize.width.should.be.eql(targetImageSize.width)
        ouputImageSize.height.should.be.eql(targetImageSize.height)

        fs.writeFileSync('out.docx', result.content)
      })

      it('image with custom size (width set and height automatic - keep aspect ratio)', async () => {
        const docxBuf = fs.readFileSync(path.join(__dirname, unit === 'cm' ? 'image-custom-size-width.docx' : 'image-custom-size-width-px.docx'))

        const targetImageSize = {
          // 2cm defined in the docx
          width: unit === 'cm' ? cmToEMU(2) : pxToEMU(100),
          // height is calculated automatically based on aspect ratio of image
          height: unit === 'cm' ? cmToEMU(0.5142851308524194) : pxToEMU(25.714330708661418)
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
            src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
          }
        })

        const ouputImageSize = await getImageSize(result.content)

        ouputImageSize.width.should.be.eql(targetImageSize.width)
        ouputImageSize.height.should.be.eql(targetImageSize.height)

        fs.writeFileSync('out.docx', result.content)
      })

      it('image with custom size (height set and width automatic - keep aspect ratio)', async () => {
        const docxBuf = fs.readFileSync(path.join(__dirname, unit === 'cm' ? 'image-custom-size-height.docx' : 'image-custom-size-height-px.docx'))

        const targetImageSize = {
          // width is calculated automatically based on aspect ratio of image
          width: unit === 'cm' ? cmToEMU(7.777781879962101) : pxToEMU(194.4444094488189),
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
            src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
          }
        })

        const ouputImageSize = await getImageSize(result.content)

        ouputImageSize.width.should.be.eql(targetImageSize.width)
        ouputImageSize.height.should.be.eql(targetImageSize.height)

        fs.writeFileSync('out.docx', result.content)
      })
    })
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

  it('image error message when width not valid param', async () => {
    return reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image-with-wrong-width.docx'))
          }
        }
      },
      data: {
        src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
      }
    }).should.be.rejectedWith(/docxImage helper requires width parameter to be valid number with unit/)
  })

  it('image error message when height not valid param', async () => {
    return reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'image-with-wrong-height.docx'))
          }
        }
      },
      data: {
        src: 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'image.png')).toString('base64')
      }
    }).should.be.rejectedWith(/docxImage helper requires height parameter to be valid number with unit/)
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

  it('input form control', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'form-control-input.docx'))
          }
        }
      },
      data: {
        name: 'Erick'
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())

    doc.getElementsByTagName('w:textInput')[0].getElementsByTagName('w:default')[0].getAttribute('w:val').should.be.eql('Erick')
  })

  it('checkbox form control', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'form-control-checkbox.docx'))
          }
        }
      },
      data: {
        ready: true
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())

    doc.getElementsByTagName('w:checkBox')[0].getElementsByTagName('w:default')[0].getAttribute('w:val').should.be.eql('1')
  })

  it('dropdown form control', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'form-control-dropdown.docx'))
          }
        }
      },
      data: {
        items: ['Boris', 'Jan', 'Barry']
      }
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())

    const entries = doc.getElementsByTagName('w:ddList')[0].getElementsByTagName('w:listEntry')

    entries.length.should.be.eql(3)
    entries[0].getAttribute('w:val').should.be.eql('Boris')
    entries[1].getAttribute('w:val').should.be.eql('Jan')
    entries[2].getAttribute('w:val').should.be.eql('Barry')
  })

  it('page break in single paragraph', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'page-break-single-paragraph.docx'))
          }
        }
      },
      data: {}
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())

    const paragraphNodes = nodeListToArray(doc.getElementsByTagName('w:p'))

    paragraphNodes[0].getElementsByTagName('w:t')[0].textContent.should.be.eql('Demo')
    paragraphNodes[1].getElementsByTagName('w:br').should.have.length(1)
    paragraphNodes[2].getElementsByTagName('w:t')[0].textContent.should.be.eql('break')
  })

  it('page break between paragraphs', async () => {
    const result = await reporter.render({
      template: {
        engine: 'handlebars',
        recipe: 'docx',
        docx: {
          templateAsset: {
            content: fs.readFileSync(path.join(__dirname, 'page-break-between-paragraphs.docx'))
          }
        }
      },
      data: {}
    })

    fs.writeFileSync('out.docx', result.content)

    const files = await decompress()(result.content)
    const doc = new DOMParser().parseFromString(files.find(f => f.path === 'word/document.xml').data.toString())

    const paragraphNodes = nodeListToArray(doc.getElementsByTagName('w:p')).filter((p) => {
      const breakNodes = getBreaks(p)

      const hasText = getText(p) != null && getText(p) !== ''

      if (!hasText && breakNodes.length === 0) {
        return false
      }

      return true
    })

    function getText (p) {
      const textNodes = nodeListToArray(p.getElementsByTagName('w:t')).filter((t) => {
        return t.textContent != null && t.textContent !== ''
      })

      return textNodes.map((t) => t.textContent).join('')
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
