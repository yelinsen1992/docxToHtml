const fs = require('fs')
const path = require('path')
const AdmZip = require('adm-zip')
const { toHTML, toPDF } = require('./index')

;(async() => {
  const docxPath = path.join(__dirname, 'test.docx')
  const targetPath = path.join(__dirname, 'result')
  const zip = new AdmZip(docxPath)
  zip.extractAllTo(targetPath, true)
  const docxBuf = fs.readFileSync(docxPath)
  await toPDF(docxBuf)
})()
