const mammoth = require('mammoth')
const fs = require('fs/promises')
mammoth
  .convertToHtml({ path: 'myFile.docx' })
  .then(function (result) {
    var html = result.value // The generated HTML
    fs.writeFile('built.html', html).catch(err => {
      console.log('html构造出错了')
    })
  })
  .catch(err => {
    console.log('转换html出错了')
  })
