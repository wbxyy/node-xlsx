const xlsx = require('node-xlsx').default
const fs = require('fs/promises')
const { readFileSync } = require('fs')
const AdmZip = require('adm-zip')

exports.readXlsx = function (filePath) {
  return xlsx.parse(readFileSync(filePath))
}

exports.buildXlsx = async function (filePath, data) {
  const newBuffer = xlsx.build(data)
  return await fs.writeFile(filePath, newBuffer)
}

exports.readDocx = function (filePath, splitMode = false) {
  const zip = new AdmZip(filePath)
  let contentXml = zip.readAsText('word/document.xml')
  // 每个段落使用数组进行分隔
  let result = []
  // 找到段落标签
  const reg_newline = /<w:p>(.*?)<\/w:p>/gis
  // 找到文本标签
  const reg_text = /<w:t>(.*?)<\/w:t>/gis
  // 找到浮动盒子的标签
  const reg_floatText = /<w:txbxContent>(.*?)<\/w:txbxContent>/gis

  // 匹配浮动盒子并删除
  contentXml = contentXml.replace(reg_floatText, '')

  // 以段落标签为分割，拆分文本
  result = [...contentXml.matchAll(reg_newline)].map(matchSectionArr => matchSectionArr[1].match(reg_text)?.reduce((pre, cur) => pre + cur.slice(5, -6), ''))

  // 对相邻的段落进行组合，分组标准为相邻之间没有空串元素

  if (splitMode) {
    //? 未完成，reduce 后，数组开头会存在 Nullish
    return result.reduce((pre, val, idx, arr) => {
      pre = [].concat(pre)
      let last = pre.length - 1

      // 判断 val 是不是空串
      if (val) {
        // 不是空串，且 pre 中有空串承接
        pre[last] = (pre[last] ?? '') + val + '\n'
      } else if (arr[idx + 1]) {
        // 当前元素是空串，检视下一个元素是否为
        pre.length++
      }
      return pre
    })
  }

  if (!splitMode) {
    return result.filter(item => !!item).join('\n')
  }
}
