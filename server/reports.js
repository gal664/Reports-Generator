const express = require("express")
const router = express.Router()
const fs = require("fs")
const path = require("path")
const tempFilesDir = "./tempFiles"
const multer = require("multer")
const upload = multer({ dest: tempFilesDir })
const xlsx = require("./utils/excel4node")

router.post("/", upload.array("upload", 3), (req, res) => {

  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  let csvData = requestFilesToArrays(req.files)

  let reportName = "Book1"
  let fileName = `${reportName}.xlsx`
  let wb = xlsx.assessmentReport(csvData[0][0][0])
  wb.write(fileName, res)
  deleteAllTempFiles()
})

function requestFilesToArrays(files) {
  filesObject = []

  for (file of files) {
    let fileObject = []
    let decoded = fs.readFileSync(path.join(__dirname, file.path), { encoding: "ucs2" })
    let decodedSplit = decoded.split("\r\n")

    for (let i = 0; i < decodedSplit.length; i++) {
      let splitRow = decodedSplit[i].split("\t")
      if (splitRow.length > 0 && splitRow[0] !== "") fileObject.push(splitRow)
    }
    filesObject.push(fileObject)
  }

  return filesObject
}

function deleteAllTempFiles() {
  fs.readdir(tempFilesDir, (err, files) => {
    if (err) throw err
    for (const file of files) {
      fs.unlink(path.join(tempFilesDir, file), err => {
        if (err) throw err
      })
    }
  })
}

module.exports = router