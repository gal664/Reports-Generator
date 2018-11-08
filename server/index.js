const express = require("express")
const router = express.Router()
const fs = require("fs")
const path = require("path")
const multer = require("multer")
const upload = multer({ dest: "tempFiles/" })
const xlsx = require("./utils/excel4node")
const parser = require("./utils/parser")
const tempFilesDir = "./tempFiles"

router.post("/", upload.array("upload", 3), (req, res, next) => {

  // get all the files in an array
  let fileData = []
  req.files.forEach(file => {
    file.path = path.join(__dirname, file.path)
    fs.readFileSync(file.path, { encoding: 'utf-8' }, (err, data) => {
      let jsonData = parser.csvToJson(data)
      fileData.push(jsonData)
    })
  })

  // create report excel file from JSON
  let reportName = "Book1"
  let fileName = `${reportName}.xlsx`
  let wb = xlsx.assessmentReport(req.files)

  // send ready excel file as response
  // wb.write(fileName, res)
  res.send(fileData)
  // delete all files in tempFiles
  fs.readdir(tempFilesDir, (err, files) => {
    if (err) throw err
    for (const file of files) {
      fs.unlink(path.join(tempFilesDir, file), err => {
        if (err) throw err
      })
    }
  })
})

module.exports = router