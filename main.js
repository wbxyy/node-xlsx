const mammoth = require('mammoth')
const util = require('./util')
const { nanoid } = require('nanoid')
const axios = require('axios')
const { requestInterceptorSignature } = require('./signature')
const config = require('./config.js')

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
  const filename = `${id}.${ext}`
  return image
    .read()
    .then(buffer => {
      return uploadFile({
        method: 'put',
        url: `/${filename}`,
        data: buffer,
      })
    })
    .then(res => {
      console.log(`图片：${filename} 已经上传成功！！`)
      return {
        src: `${config.COS_BASE_URL}/${filename}`,
      }
    })

  //! 直接将 image 对象作为 data 上传至 COS 服务（不行，无法通过 url 访问图片）
  //! 获取 base64 作为 data 上传至 COS 服务（不行，无法通过 url 访问图片）
})

const { readdir } = require('fs/promises')
const { resolve, join } = require('path')
;(async () => {
  // 拿到文件夹绝对路径
  const dir = resolve(__dirname, 'docx')

  // 拿到文件名数组
  const paths = await readdir(dir)

  const responses = await Promise.all(paths.map(path => mammoth.convertToHtml({ path: join(dir, path) }, { convertImage: imgConvert })))

  const sheetData = responses.map((res, index) => [paths[index].replace('.docx', ''), '仓储', res.value])

  const data = [{ name: 'sheet1', data: sheetData }]
  util.buildXlsx('built.xlsx', data)
})()

// 类型侦察

// 关键字：仓储、仓库、库存、库位、云仓、进仓、入仓，总览，出仓，结余，作业量，周转率、利用率、作业、代开证、订单、统计、冻结、担保、货权，货权转移、货转、预约、质押、质物、质权、出质、仓租、散租、整租、司机、合同、对账单、运输、代售、物流、配送
