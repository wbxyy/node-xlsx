const xlsx = require('node-xlsx').default
const fs = require('fs/promises')
const { readFileSync } = require('fs')

//! parse a xlsx file
const buffer = xlsx.parse(readFileSync(`${__dirname}/myFile.xlsx`))
console.log(typeof buffer)

// xlsx 转 sheets 数组，遍历数据
// buffer.forEach(sheet => {
//   console.log(sheet.name)
//   sheet.data.forEach(item => {
//     console.log(item)
//   })
// })

//! build a xlsx file
const row_1 = ['担保', '货物担保的录入、录出、统计、估价的描述', '1、估价页中的分类不按客户现存货物统计，而是每次录入担保时，根据录入的货物品牌规格来更新估价页记录。\n' + '2、估价页每条分类记录的件数、重量、总估价只统计已录入到担保的货物。\n' + '3、当估价页的某分类的货物全部录出时，该分类记录仍然保留当前使用价格和备查价格，但相关统计值如件数、重量、总估价应实时变更为0。']
const row_2 = ['仓储', '对出仓单据进行添加出仓时，不能选择某些现存货物的bug']
const data = [row_1, row_2]

//todo: returns a buffer
const newBuffer = xlsx.build([{ name: '测试', data }])
console.log(newBuffer)

//todo: write a file

fs.writeFile('built.xlsx', newBuffer).catch(err => {
  console.log('出错了，原因如下：')
  console.log(err)
})

//!  读取 word 文件的内容会乱码吗，👎
//! 会的

// fs.readFile('myFile.docx').then(res => {
//   console.log(res.toString())
// })

//! 分析得知，docx 文件的本质是压缩包，其中是对 xml 的封装
//! 从此完成解析步骤：
//todo: zip 解压缩
//todo: 读取内部 xml 文件作为文本
//todo: 使用正则表达式去除标签

const AdmZip = require('adm-zip')
//* filePath 为文件路径
const zip = new AdmZip('myFile.docx')
let contentXml = zip.readAsText('word/document.xml')
const result = []
contentXml.match(/<w:t>(.*?)<\/w:t>/gis).forEach(item => {
  result.push(item.slice(5, -6))
})

console.log(result)

//! 将步骤封装函数
function readTextDocx(filePath) {
  const zip = new AdmZip(filePath)
  let contentXml = zip.readAsText('word/document.xml')
  console.log(contentXml)
  const result = []
  contentXml.match(/<w:t>(.*?)<\/w:t>/gis).forEach(item => {
    result.push(item.slice(5, -6))
  })
  return result
}

//! 不足之处
//todo 只能提取文本内容
//todo 文本内容没有断句逻辑
const result2 = readTextDocx('myFile2.docx')
console.log(result2)
