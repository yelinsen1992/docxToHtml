const fs = require('fs')
const path = require('path')
const awaitWrap = (promise) => {
  return promise
    .then(data => [null, data])
    .catch(err => [err, null])
}
// 递归删除
const rmDir = (dir) => {
  return new Promise(resolve => {
    fs.stat(dir, (err, status) => {
      if (err) {
        resolve(false)
      } else {
        if (status.isDirectory()) { // 是文件夹
          fs.readdir(dir, (err, file) => {
            if (!err) {
              const res = file.map((item) => rmDir(path.join(dir, item)))
              Promise.all(res).then(() => { // 当所有的子文件都删除后就删除当前文件夹
                fs.rmdir(dir, () => { resolve(true) })
              })
            } else {
              resolve(false)
            }
          })
        } else {
          fs.unlink(dir, () => { resolve(true) })
        }
      }
    })
  })
}
// 读取路径信息
const getStat = (path) => {
  return new Promise(resolve => {
    fs.stat(path, (err, stats) => {
      err ? resolve(false) : resolve(stats)
    })
  })
}
// 创建路径
const mkdir = (dir) => {
  return new Promise(resolve => {
    fs.mkdir(dir, err => {
      err ? resolve(false) : resolve(true)
    })
  })
}
// 递归判断上级目录
const mkDir = async(dir) => {
  const isExists = await getStat(dir)
  if (isExists && isExists.isDirectory()) { // 存在路径且是文件夹，返回true
    return true
  } else if (isExists) {
    return false // 存在路径且是文件，返回false
  }
  const tempDir = path.parse(dir).dir // 拿到上级路径
  const status = await mkDir(tempDir) // 递归判断，如果上级目录也不存在，则会代码会在此处继续循环执行，直到目录存在
  let mkdirStatus
  if (status) {
    mkdirStatus = await mkdir(dir)
  }
  return mkdirStatus
}
// 创建文件
const writeFile = async(dir, data) => {
  return new Promise(resolve => {
    fs.writeFile(dir, data, err => {
      err ? resolve(false) : resolve(true)
    })
  })
}
// 删除文件
const unLink = async(dir) => {
  return new Promise(resolve => {
    fs.unlink(dir, err => {
      err ? resolve(false) : resolve(true)
    })
  })
}
const readFile = async(dir) => {
  return new Promise(resolve => {
    fs.readFile(dir, (err, result) => {
      err ? resolve(false) : resolve(result)
    })
  })
}
module.exports = {
  awaitWrap,
  mkDir,
  rmDir,
  writeFile,
  unLink,
  readFile
}
