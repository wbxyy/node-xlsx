const axios = require('axios')
const { readFile } = require('fs/promises')
const { requestInterceptorSignature } = require('../signature')
const config = require('../config.js')
const uploadFile = axios.create({
  baseURL: config.COS_BASE_URL,
})
// axios.defaults.baseURL = config.COS_BASE_URL
//! 使用第三方服务签名接口时，一定要把 axios 默认的 headers 置空(原因不明)
uploadFile.defaults.headers = {}
uploadFile.interceptors.request.use(requestInterceptorSignature)
const filePath = 'img4.png'
readFile(filePath).then(data => {
  uploadFile({
    method: 'put',
    url: `/${filePath}`,
    data,
  }).then(
    res => {
      console.log(res)
    },
    err => {
      console.log(err)
      console.log(err?.response.status + err?.response.statusText)
    }
  )
})
