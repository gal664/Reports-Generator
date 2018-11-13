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
  // console.log(csvData)
  
  let reportName = "Book1"
  let fileName = `${reportName}.xlsx`
  let wb = xlsx.assessmentReport(csvData[0].data[0][0]) // example for putting the first item of first array into cell A1 in the first sheet
  wb.write(fileName, res)
  deleteAllTempFiles()
})

function requestFilesToArrays(files) {
  let filesArray = []

  for (file of files) {
    let fileObject = {
      data: [],
      name: file.originalname.slice(0, file.originalname.indexOf("crosstab")).replace(/_/g, " ").trim(),
    }
    let decoded = fs.readFileSync(path.join(__dirname, file.path), { encoding: "ucs2" })
    let decodedSplit = decoded.split("\r\n")

    for (let i = 0; i < decodedSplit.length; i++) {
      let splitRow = decodedSplit[i].split("\t")
      if (splitRow.length > 0 && splitRow[0] !== "") fileObject.data.push(splitRow)
    }
    filesArray.push(fileObject)
  }

  return filesArray
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