
module.exports = (files) => {
  const f = files.find(f => f.path === 'word/document.xml')

  // this can get much more complex... check the previous boris implementation
  f.data = f.data.toString().replace(/<w:t><docxPageBreak \/><\/w:t>/g, (val) => {
    return '<w:br w:type="page"/>'
  })
}
