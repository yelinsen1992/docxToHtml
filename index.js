const path = require('path')
const AdmZip = require('adm-zip')
const {
  mkDir,
  writeFile,
  unLink,
  readFile
} = require('./fuc.js')
const puppeteer = require('puppeteer')
const outDir = path.join(__dirname, 'tmp')
// 时间命名文件
const getTime = () => {
  return (new Date()).getTime()
}
// 对象去重，以第一个为基准保留
const removeDup = (objA, objB) => {
  const objNew = {}
  for (const keyA in objA) {
    for (const keyB in objB) {
      if (keyA !== keyB) {
        objNew[keyA] = objA[keyA]
        objNew[keyB] = objB[keyB]
      } else if (keyA === keyB) {
        objNew[keyA] = objA[keyA]
      }
    }
  }
  return objNew
}
// 设置html行内样式
const setStyle = (item) => {
  let text = ''
  if (JSON.stringify(item) !== '{}') {
    for (const i in item) {
      text += i.replace('_', '-') + ':' + item[i] + ';'
    }
  }
  return text
}
// 将docx样式转换为html样式
const getStyle = (item) => {
  const obj = {}
  const lineHeight = item.match(/w:line="(.*?)"/gi) // 行距
  const textAlign = item.match(/w:jc w:val="(.*?)"\/>/gi) // 段落对齐方式
  const fontWeight = item.match(/(<w:b\/>)|(<w:bCs\/>)/gi) // 加粗
  const fontStyle = item.match(/(<w:i\/>)|(<w:iCs\/>)/gi) // 斜体
  const fontSize = item.match(/<w:sz w:val="(.*?)"\/>/gi) // 字体大小
  const fontColor = item.match(/<w:color w:val="(.*?)"/gi) // 字体颜色
  const fontFamily = item.match(/w:ascii="(.*?)"/gi) // 字体
  const marginTop = item.match(/w:spacing(.*?)w:before="(.*?)"(.*?)/)
  const marginBottom = item.match(/w:spacing(.*?)w:after="(.*?)"(.*?)/)
  if (lineHeight) { obj.line_height = lineHeight[0].slice(8, -1) / 2.4 + '%' }
  if (textAlign) { obj.text_align = (textAlign[0].slice(12, -3) === 'distribut' || textAlign[0].slice(12, -3) === 'both' ? 'justify' : textAlign[0].slice(12, -3)) }
  if (fontWeight) { obj.font_weight = 'bold' }
  if (fontStyle) { obj.font_style = 'italic' }
  if (fontSize) { obj.font_size = fontSize[0].slice(13, -3) / 2 + 'pt' }
  if (fontColor && fontColor[0].slice(16, -1) !== 'auto') { obj.color = '#' + fontColor[0].slice(16, -1) }
  if (marginTop) { obj.margin_top = marginTop[2] / 20 + 'pt' }
  if (marginBottom) { obj.margin_bottom = marginBottom[2] / 20 + 'pt' }
  if (fontFamily) { obj.font_family = fontFamily[0].slice(9, -1) }
  return obj
}
// 解析docx图片映射，返回rId映射图片地址/图片base64
const parserPic = async (picXml, zip) => {
  return new Promise((resolve) => {
    const mR = picXml.match(/<Relationship Id="(.*?)"(.*?)\/>/gi)
    const picObj = {}
    if (mR) {
      mR.forEach(item => {
        if (item.indexOf('image') !== -1) {
          const rId = item.match(/Id="(.*?)"/)[1]
          const fileName = item.match(/media\/(.*?)"/)[1]
          const entry = zip.getEntry('word/media/' + fileName)
          const buffer = entry.getData()
          const base64Str = 'data:' + fileName.slice(fileName.indexOf('.') + 1) + ';base64,' + buffer.toString('base64')
          picObj[rId] = base64Str
        }
      })
    }
    resolve(picObj)
  })
}
// 解析docx默认标题、默认段落样式
const parserStyle = (stylesXml) => {
  const mWStyle = stylesXml.match(/<w:style w:type="paragraph"(.*?)w:styleId="(.*?)">(.*?)<\/w:style>/gi)
  const mWDOC = stylesXml.match(/<w:docDefaults>(.*?)<\/w:docDefaults>/gi)
  const mPDEFAULT = stylesXml.match(/<w:style w:type="paragraph" w:default="1"(.*?)<\/w:style>/gi)
  const defaultObj = {}
  let defaultObj2 = {}
  let fontObj = ''
  if (mWDOC) {
    const defaultFont = mWDOC[0].match(/w:eastAsia="(.*?)"/)
    if (defaultFont) {
      fontObj = defaultObj.font_family = defaultFont[1]
    }
  }
  if (mPDEFAULT) {
    defaultObj2 = getStyle(mPDEFAULT[0])
    let mT = mPDEFAULT[0].match(/w:spacing(.*?)w:before="(.*?)"(.*?)/)
    let mB = mPDEFAULT[0].match(/w:spacing(.*?)w:after="(.*?)"(.*?)/)
    let mL = mPDEFAULT[0].match(/w:spacing(.*?)w:left="(.*?)"(.*?)/)
    let mR = mPDEFAULT[0].match(/w:spacing(.*?)w:right="(.*?)"(.*?)/)
    mT = mT ? mT[2] / 20 + 'pt' : 0
    mB = mB ? mB[2] / 20 + 'pt' : 0
    mL = mL ? mL[2] / 20 + 'pt' : 0
    mR = mR ? mR[2] / 20 + 'pt' : 0
    defaultObj2.margin = mT + ' ' + mR + ' ' + mB + ' ' + mL
  }
  const normal = '.normal{' + setStyle(removeDup(defaultObj2, defaultObj)) + '}'
  const obj = {}
  if (mWStyle) {
    mWStyle.forEach(item => {
      const styleId = item.match(/styleId="(.*?)"/)[1]
      obj[styleId] = getStyle(item)
      if (fontObj) {
        obj[styleId].font_family = fontObj
      }
    })
  }
  let text = ''
  for (const key in obj) {
    text += 'h' + key + '{' + setStyle(obj[key]) + '}\r\n'
  }
  return text + normal
}
// 解析docx内容，返回拼接好的html页面
const toHTML = async (docxBuf, name) => {
  const tmpName = getTime()
  const docxTmp = path.join(outDir, tmpName + '.docx')
  await mkDir(outDir) // 新建目录
  await writeFile(docxTmp, docxBuf)
  const zip = new AdmZip(docxTmp)
  const contentXml = zip.readAsText('word/document.xml') // 将document.xml(解压缩后得到的文件)读取为text内容
  const picXml = zip.readAsText('word/_rels/document.xml.rels')
  const stylesXml = zip.readAsText('word/styles.xml')
  const picObj = await parserPic(picXml, zip, outDir) // 保存docx上图片的id映射信息，包括图片名字和图片二进制数据
  const defaultStyle = parserStyle(stylesXml) // 默认样式
  const mWP = contentXml.match(/(<w:p>(.*?)<\/w:p>)|(<w:p\/>)/gi) // match匹配每个段落
  const parserJson = []
  let html = ''
  if (mWP) {
    mWP.forEach((wpItem, i) => { // 段落循环
      let className = ' class="normal"'
      if (wpItem.indexOf('<w:p/>') === -1) {
        const wpObj = {}
        const mWPPR = wpItem.match(/<w:pPr>(.*?)<\/w:pPr>/gi)[0] // match匹配每个段落段落样式内容块
        const mWR = wpItem.match(/<w:r>(.*?)<\/w:r>/gi) // match匹配每个段落中每个文本样式串块
        const styleJson = getStyle(mWPPR)
        const styleText = JSON.stringify(styleJson) === '{}' ? '' : 'style="' + setStyle(styleJson) + '"'
        let label = 'p'
        const styleId = wpItem.match(/<w:pStyle w:val="(.*?)"\/>/)
        if (styleId) {
          label = 'h' + styleId[1]
          className = ''
        }
        wpObj.styleText = styleJson
        html += '<' + label + className + ' ' + styleText + '>'
        wpObj.wr = []
        if (mWR) {
          mWR.forEach((wrItem, j) => { // 段落中文字块循环
            const wrObj = {}
            const styleJson = getStyle(wrItem)
            const styleText = JSON.stringify(styleJson) === '{}' ? '' : 'style="' + setStyle(styleJson) + '"'
            const mWT = wrItem.match(/(<w:t>.*?<\/w:t>)|(<w:t\s.[^>]*?>.*?<\/w:t>)/gi) // 获取文字
            const mPIC = wrItem.match(/<w:drawing>(.*?)<\/w:drawing>/gi)
            if (mWT) {
              wrObj.text = mWT[0].indexOf('xml:space') === -1 ? mWT[0].slice(5, -6) : mWT[0].slice(26, -6)
              wrObj.styleText = styleJson
              html += '<span ' + styleText + '>' + wrObj.text + '</span>'
            }
            if (mPIC) {
              wrObj.styleText = styleJson
              const picSize = mPIC[0].match(/<a:ext cx="(.*?)" cy="(.*?)"\/>/)
              const width = picSize[1] / 1440 / 7 + 'pt'
              const height = picSize[2] / 1440 / 7 + 'pt'
              const title = mPIC[0].match(/<pic:cNvPr(.*?)descr="(.*?)"\/>/)
              const rId = mPIC[0].match(/<a:blip r:embed="(.*?)"\/>/)[1]
              const obj = {}
              if (width) { obj.width = width }
              if (height) { obj.height = height }
              if (title) { obj.title = title[2] }
              if (rId) { obj.src = picObj[rId] }
              html += '<span ' + styleText + '><img width="' + width + '" height="' + height + '" title="' + (title || '') + '" src="' + picObj[rId] + '"></span>'
            }
            wpObj.wr.push(wrObj)
          })
        } else {
          html += '<span ' + styleText + '>&nbsp;</span>'
        }
        parserJson.push(wpObj)
        html += '</' + label + '>'
      } else {
        html += '<p ' + className + '><span>&nbsp;</span></p>'
      }
    })
  }
  const htmlFront = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<meta http-equiv="X-UA-Compatible" content="IE=edge">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>' + (name || getTime()) + '</title>' +
    '</head>' +
    '<style>' + defaultStyle + '</style>' +
    '<body><div style="page-break-after:always;">'
  const htmlRear = '</div></body></html>'
  const documentHtml = htmlFront + html + htmlRear
  await unLink(docxTmp)
  return documentHtml
}
// 解析docx内容，返回pdf文件的buffer
const toPDF = async (docxBuf, name) => {
  const html = await toHTML(docxBuf, name)
  const tmpName = getTime()
  const htmlTmp = path.join(outDir, tmpName + '.html')
  const pdfTmp = path.join(outDir, tmpName + '.pdf')
  await writeFile(htmlTmp, html)
  const browser = await puppeteer.launch()
  const page = await browser.newPage()
  await page.goto(htmlTmp, { waitUntil: 'networkidle2' })
  await page.pdf({
    path: pdfTmp,
    format: 'a4',
    margin: {
      top: 72,
      right: 90,
      bottom: 72,
      left: 72
    }
  })
  await browser.close()
  const pdfBuf = await readFile(pdfTmp)
  await unLink(htmlTmp)
  await unLink(pdfTmp)
  return pdfBuf
}
module.exports = {
  toHTML,
  toPDF
}
