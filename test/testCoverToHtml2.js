const util = require('../util')

// 读取 xlsx 文件
const xlsxObj = util.readXlsx('myFile.xlsx')

// 遍历 xlsxObj 的内容
// xlsxObj.forEach(sheet => {
//   sheet.data.forEach(item => {
//     console.log(item)
//   })
// })

// 构建 xlsx 文件
const data = [
  {
    name: '清单',
    data: [
      ['仓储', '仓储一堆巴拉巴拉'],
      ['代开证', '代开证一堆巴拉巴拉'],
    ],
  },
  {
    name: '未完成',
    data: [
      ['仓储', '仓储一堆未完成巴拉巴拉'],
      ['代开证', '代开证一堆未完成巴拉巴拉'],
    ],
  },
]

util.buildXlsx('built.xlsx', data).catch(err => {
  console.log('构建 xlsx 出错')
})

// const docx1 = util.readDocx('myFile.docx')
// const docx2 = util.readDocx('myFile2.docx')
// const docx3 = util.readDocx('myFile3.docx')
// const docx4 = util.readDocx('myFile4.docx')
// const docx5 = util.readDocx('myFile5.docx')

// console.log(docx1)
// console.log(docx2)
// console.log(docx3)
// console.log(docx4)
// console.log(docx5)

//! 好了，上面是 api 热身，下面是干活
//todo: 读取目录中的 docx
//todo: docx 逐个解析，获得字符串数组
//todo: 构造 json 格式的对象，再构造 xlsx 文件
const { readdir } = require('fs/promises')
const { resolve, join } = require('path')
function buildXlsx() {
  const dir = resolve(__dirname, 'docx')

  readdir(dir)
    .then(res => {
      const result = []
      res.forEach(path => {
        path = result.push([path.replace('.docx', ''), '仓储', util.readDocx(join(dir, path))])
      })
      return result
    })
    .then(entries => {
      const data = [{ name: 'sheet1', data: entries }]
      util.buildXlsx('built.xlsx', data)
    })
}

//! 新思路
//todo: 使用 mammoth 将 docx 转换为 html 文件
//todo: 使用 html 格式的文本构造 xlsx
const mammoth = require('mammoth')
const fs = require('fs/promises')
const { nanoid } = require('nanoid')
// 解析 docx 的图片转换器
// 返回一个 带 src 属性的对象，这个 src 会被嵌入到生成的 img 标签中
const imgConvert = mammoth.images.imgElement(function (image) {
  // return {
  //   src: './docx/abc.png',
  // }
  return image.read('base64').then(function (imageBuffer) {
    const id = nanoid()
    return {
      // src: 'data:' + image.contentType + ';base64,' + imageBuffer,
      src: `haha.png`,
    }
  })
})
console.log(imgConvert)
mammoth.convertToHtml({ path: 'myFile.docx' }, { convertImage: imgConvert }).then(ret => {
  const html = ret.value
  fs.writeFile('built.html', html).catch(err => {
    console.log('构建html出错')
  })
  // const entries = [['myFile', '仓储', html]]
  // const data = [{ name: 'sheet1', data: entries }]
  // util.buildXlsx('built.xlsx', data)
})
