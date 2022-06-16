const crypto = require('crypto')
const moment = require('moment')
const globalConfigFile = require('./config')

exports.requestInterceptorSignature = function (config) {
  //! keyTime(✔)
  const unix = moment().unix()
  const expire = 3600 // 1小时
  const keyTime = `${+unix};${+unix + expire}`

  //! 生成 UrlParamList 和 HttpParameters(✔)
  function generator(src) {
    src = { ...src }
    const map = {}
    for (key in src) {
      encodeKey = encodeURIComponent(key.toLowerCase())
      encodeVal = encodeURIComponent(src[key])
      map[encodeKey] = encodeVal
    }
    const keyList = Object.keys(map).sort()
    let [parameters, list] = ['', '']
    if (keyList) {
      parameters = keyList.map(key => `${key}=${map[key]}`).join('&')
      list = keyList.join(';')
    }
    return [parameters, list]
  }

  let [httpParameters, urlParamList] = generator(config.params)

  //! 生成 HeaderList 和 HttpHeaders(✔)
  const headersCustomer = {
    // 'User-Agent': 'axios/0.27.2',
    // Connection: 'close',
    // 'Content-Length': 0,
  }
  for (key in config.headers) {
    if (typeof config.headers[key] === 'string') {
      headersCustomer[key] = config.headers[key]
    }
  }
  const method = config.method
  const header = { ...config.headers['common'], ...config.headers[method], ...headersCustomer }

  let [HttpHeaders, HeaderList] = generator(header)

  //! 生成 HttpString(✔)
  const httpString = `${config.method.toLowerCase()}\n${config.url}\n${httpParameters}\n${HttpHeaders}\n`

  //! 生成 stringToSign
  const sha1HttpString = crypto.createHash('sha1').update(httpString).digest('hex')
  const stringToSign = `sha1\n${keyTime}\n${sha1HttpString}\n`

  //! 生成 signKey
  //! 很重要的，不能泄露的 密钥
  const yourSecretKey = globalConfigFile.COS_SECRET_KEY

  const signKey = crypto.createHmac('sha1', yourSecretKey).update(keyTime).digest('hex')

  //! 生成 Signature
  const signature = crypto.createHmac('sha1', signKey).update(stringToSign).digest('hex')

  //! 生成签名
  const yourSecretId = globalConfigFile.COS_SECRET_ID
  const signStringResult = ['q-sign-algorithm=sha1', `q-ak=${yourSecretId}`, `q-sign-time=${keyTime}`, `q-key-time=${keyTime}`, `q-header-list=${HeaderList}`, `q-url-param-list=${urlParamList}`, `q-signature=${signature}`].join('&')
  config.headers['Authorization'] = signStringResult

  return config
}
