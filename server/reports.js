const express = require("express")
const router = express.Router()
const fs = require("fs")
const path = require("path")
const tempFilesDir = "./tempFiles"
const multer = require("multer")
const upload = multer({ dest: tempFilesDir })
const xlsx = require("./utils/excel4node")

let uploadInputs = upload.fields([{ name: 'grades_by_subject', maxCount: 1 }, { name: 'struggling_students', maxCount: 1 }, { name: 'grades_by_question', maxCount: 1 }])

router.post("/assessmentReport", uploadInputs, (req, res) => {
  
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }
  
  let wb = xlsx.assessmentReport(req.query, requestFilesToArrays(req.files))
  wb.write(`${req.query.classes} - ${req.query.assessmentname} - דוח מבחן.xlsx`, res)
  deleteAllTempFiles()
  
})

// router.post("/practiceReport", uploadInputs, (req, res) => {
  
//   if (!req.files) {
//     reject("no files uploaded")
//     res.status(500).send("error uploading the files")
//   }
  
//   let wb = xlsx.assessmentReport(req.query, requestFilesToArrays(req.files))
//   wb.write(`${req.query.classes} - ${req.query.assessmentname} - דוח מבחן.xlsx`, res)
//   deleteAllTempFiles()
  
// })




function requestFilesToArrays(files) {
  let filesArray = []

  Object.keys(files).forEach(key => {
    let fileObject = { data: [], name: key }
    let decoded = fs.readFileSync(path.join(__dirname, files[key][0].path), { encoding: "ucs2" })
    let decodedSplit = decoded.split("\r\n")

    for (let i = 0; i < decodedSplit.length; i++) {
      let splitRow = decodedSplit[i].split("\t")
      if (splitRow.length > 0 && splitRow[0] !== "") fileObject.data.push(splitRow)
    }
    filesArray.push(fileObject)
  });

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