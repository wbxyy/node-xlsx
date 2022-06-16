const mammoth = require('mammoth')
const util = require('../util')
const fs = require('fs/promises')
const { nanoid } = require('nanoid')
const axios = require('axios')
const { requestInterceptorSignature } = require('../signature')
const config = require('../config.js')

const uploadFile = axios.create({
  baseURL: config.COS_BASE_URL,
})
uploadFile.defaults.headers = {}
uploadFile.interceptors.request.use(requestInterceptorSignature)

// 通过 imgCovert 中的回调函数，将 buffer 上传至 COS 服务
const imgConvert = mammoth.images.imgElement(function (image) {
  const id = nanoid()
  //* 文档提到：如果 read 没有指定编码，则返回图像 buffer
  const extPattern = /image\/(\w+)/
  const ext = image.contentType.match(extPattern)[1]
  return image
    .read()
    .then(buffer => {
      return uploadFile({
        method: 'put',
        url: `/${id}.${ext}`,
        data: buffer,
      })
    })
    .then(res => {
      console.log(res)
      return {
        src: `${config.COS_BASE_URL}/${id}.${ext}`,
      }
    })

  //! 直接将 image 对象作为 data 上传至 COS 服务（不行，无法通过 url 访问图片）
  //! 获取 base64 作为 data 上传至 COS 服务（不行，无法通过 url 访问图片）
})

mammoth.convertToHtml({ path: 'myFile.docx' }, { convertImage: imgConvert }).then(ret => {
  const html = ret.value
  // fs.writeFile('built.html', html).catch(err => {
  //   console.log('构建html出错')
  // })
  const entries = [['myFile', '仓储', html]]
  const data = [{ name: 'sheet1', data: entries }]
  util.buildXlsx('built.xlsx', data)
})
