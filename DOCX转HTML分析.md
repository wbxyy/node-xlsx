# 分析比写代码更重要

## 本日课题
将 DOCX 转换 HTML 文档，以此为原料构建 xlsx

### 使用工具

**mammoth**
  docx 解析工具，将 docx 转换为 html 文档


### 操作分析

#### 操作 docx

    mammoth 获取 docx 的 html 格式
    util.buildXlsx 构建 xlsx


### 实现难点

**难点一：html 文档中图片的处理**

  mammoth 获取 html 文档后，图片默认以内联方式（base64） 存放于 html。

  保存 base64 的 html 文档体积很大，无法存储到 xlsx 的单元格中（单元格字符数限制 32767）

  思路：将 base64 图片进行解析，上传到腾讯云 COS 服务中存储。将 html 文档中的 base64 替换为 COS 服务的链接。

**难点二：图片的上传**

  image 对象由 mammoth 图片转换器回调函数提供，以对应 docx 中的每个图片的信息。

  **1. 分析 image 对象的成分**

  image 对象拥有 contentType 属性，read 函数

  易知 read 函数封装了流的操作，直接调用可以获得图片 buffer

  **2. 进一步尝试：**

  图片转换器回调中，用 read 获取图像 buffer，将 buffer 作为 data 上传到 COS 服务，文件名称由 nanoid 生成，文件后缀名从 contentType 提取。
  图片转换器回调返回一个带 src 属性的对象，src 值为 COS 服务的图片 url，即可改写转换后 html 的 img 标签
  最后将 html 的内容写入 xlsx

**难点三：接口签名**

  难在数据叠罗汉，中间步骤无法验算。
  按照腾讯云COS的文档构建签名：

    1. 生成keyTime
    2. 生成 UrlParamList 和 HttpParameters
    3. 生成 HeaderList 和 HttpHeaders
    4. 生成 HttpString
    5. 生成 stringToSign
    6. 生成 signKey
    7. 生成 Signature
    8. 生成签名

  步骤相当多，本质是

    1. 通过时间戳、请求参数、请求头组成校验和的原料（官方称httpString），用 httpString 创建校验和，用的是 createHash。
    2. 用 secretKey（腾讯云COS服务提供）作为密钥，keyTime 作为消息，用 sha1 生成签名密钥 signKey
    3. 用 signKey 作为密钥，校验和作为消息，用 sha1 生成签名 signature
    4. 把签名、及一堆用来生成签名的中间产物，放在请求头
