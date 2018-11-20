const fs = require("fs")
const path = require("path")

const express = require("express")
const router = express.Router()

const tempFilesDir = "./tempFiles"
const multer = require("multer")
const upload = multer({ dest: tempFilesDir })

const docx = require("docx")

const excelWorker = require("./utils/excelWorker")
const wordWorker = require("./utils/wordWorker")
const filesParser = require("./utils/filesParser")
const dirCleaner = require("./utils/dirCleaner")

// define all the inputs that multer should accept files from
let uploadInputs = upload.fields([
  { name: "grades_by_subject", maxCount: 1 },
  { name: "struggling_students", maxCount: 1 },
  { name: "grades_by_question", maxCount: 1 },
  { name: "class_grades_by_subject", maxCount: 1 },
  { name: "class_grades_by_question", maxCount: 1 },
  { name: "student_data", maxCount: 1 },
  { name: "recommendations_data", maxCount: 1 },
])

// exports an excel report for assessment data
router.post("/assessment", uploadInputs, (req, res) => {
  
  // handle cases with no files
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  // parse the request files
  let parsedReqFiles = filesParser(req.files)

  // create a workbook, passing the query string parameters and parsed file data
  let wb = excelWorker.assessmentReport(req.query, parsedReqFiles)

  // when the workbook is created, return it in the response
  wb.write(`${req.query.classes} - ${req.query.assessmentname} - דוח מבחן.xlsx`, res)
  
  // when response has been sent, clear tempFiles
  dirCleaner(tempFilesDir)
})

// exports an excel report for practice data
router.post("/practice", uploadInputs, (req, res) => {
  
  // handle cases with no files
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  // parse the request files
  let parsedReqFiles = filesParser(req.files)

  // create a workbook, passing the query string parameters and parsed file data  
  let wb = excelWorker.practiceReport(req.query, parsedReqFiles)

  // when the workbook is created, return it in the response
  wb.write(`${req.query.classes} - ${req.query.assessmentname} - דוח תרגול.xlsx`, res)

  // when response has been sent, clear tempFiles
  dirCleaner(tempFilesDir)
})

// exports an excel report for grade assessment data
router.post("/gradeAssessment", uploadInputs, (req, res) => {

  // handle cases with no files
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  // parse the request files
  let parsedReqFiles = filesParser(req.files)

  // create a workbook, passing the query string parameters and parsed file data  
  let wb = excelWorker.practiceReport(req.query, parsedReqFiles)

  // when the workbook is created, return it in the response
  wb.write(`${req.query.classes} - ${req.query.assessmentname} - דוח מבחן שכבתי.xlsx`, res)

  // when response has been sent, clear tempFiles
  dirCleaner(tempFilesDir)
})

// exports an excel report for student data
router.post("/student", uploadInputs, (req, res) => {

  // create an instance of docx packer to create a .docx file from the data that will be created
  const packer = new docx.Packer()
  
  // define the name for the file that will be sent
  const fileName = `${req.query.classes} - ${req.query.assessmentname} - דוח תלמיד.docx`
  
  // define the path for the file that will be sent
  const filePath = path.join(tempFilesDir, fileName)

  // handle cases with no files
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  // parse the request files
  let parsedReqFiles = filesParser(req.files)

  // create a doc from the parsed files' data
  let doc = wordWorker.studentReport(req.query, parsedReqFiles)

  //send the doc created in the response using packer instance defined above
  packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(filePath, buffer)
    res.download(filePath)
  })

  // when response has been sent, clear tempFiles
  dirCleaner(tempFilesDir)
})

// exports an excel report for recommendations data
router.post("/recommendations", uploadInputs, (req, res) => {

  // create an instance of docx packer to create a .docx file from the data that will be created
  const packer = new docx.Packer()

  // define the name for the file that will be sent
  const fileName = `${req.query.classes} - ${req.query.assessmentname} - דוח המלצות.docx`

  // define the path for the file that will be sent
  const filePath = path.join(tempFilesDir, fileName)

  // handle cases with no files
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  // parse the request files
  let parsedReqFiles = filesParser(req.files)

  //send the doc created in the response using packer instance defined above
  let doc = wordWorker.recommendationsReport(req.query, parsedReqFiles)
  packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(filePath, buffer)
    res.download(filePath)
  })

  // when response has been sent, clear tempFiles
  dirCleaner(tempFilesDir)
})

module.exports = router
