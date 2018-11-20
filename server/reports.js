const express = require("express");
const router = express.Router();
const fs = require("fs");
const path = require("path");
const tempFilesDir = "./tempFiles";
const multer = require("multer");
const upload = multer({ dest: tempFilesDir });
const xlsx = require("./utils/excel4node");
const docx = require("docx");
const docxFunctions = require("./utils/nodeDocx");
const filesParser = require("./utils/filesParser")
const dirCleaner = require("./utils/dirCleaner")

let uploadInputs = upload.fields([
  { name: "grades_by_subject", maxCount: 1 },
  { name: "struggling_students", maxCount: 1 },
  { name: "grades_by_question", maxCount: 1 },
  { name: "class_grades_by_subject", maxCount: 1 },
  { name: "class_grades_by_question", maxCount: 1 },
  { name: "student_data", maxCount: 1 },
  { name: "recommendations_data", maxCount: 1 },
]);

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
  let wb = xlsx.assessmentReport(req.query, parsedReqFiles)

  // when the workbook is created, return it in the response
  wb.write(`${req.query.classes} - ${req.query.assessmentname} - דוח מבחן.xlsx`, res)
  
  // when response has been sent, clear tempFiles
  dirCleaner(tempFilesDir)
});

router.post("/practice", uploadInputs, (req, res) => {
  if (!req.files) {
    reject("no files uploaded");
    res.status(500).send("error uploading the files");
  }

  let wb = xlsx.practiceReport(req.query, filesParser(req.files));
  wb.write(
    `${req.query.classes} - ${req.query.assessmentname} - דוח תרגול.xlsx`,
    res
  );

  dirCleaner(tempFilesDir)
});

router.post("/gradeAssessment", uploadInputs, (req, res) => {
  if (!req.files) {
    reject("no files uploaded");
    res.status(500).send("error uploading the files");
  }

  let wb = xlsx.gradeAssessmentReport(
    req.query,
    filesParser(req.files)
  );
  wb.write(
    `${req.query.classes} - ${req.query.assessmentname} - דוח מבחן שכבתי.xlsx`,
    res
  );

  dirCleaner(tempFilesDir)
});

router.post("/student", uploadInputs, (req, res) => {
  const packer = new docx.Packer();
  const fileName = `${req.query.classes} - ${
    req.query.assessmentname
  } - דוח תלמיד.docx`;
  const filePath = path.join(tempFilesDir, fileName);

  if (!req.files) {
    reject("no files uploaded");
    res.status(500).send("error uploading the files");
  }

  let doc = docxFunctions.studentReport(
    req.query,
    filesParser(req.files)
  );

  packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(filePath, buffer);
    res.download(filePath);
  });

  dirCleaner(tempFilesDir)
});

router.post("/recommendations", uploadInputs, (req, res) => {
  const packer = new docx.Packer()
  const fileName = `${req.query.classes} - ${req.query.assessmentname} - דוח המלצות.docx`
  const filePath = path.join(tempFilesDir, fileName)
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  let doc = docxFunctions.recommendationsReport(req.query, filesParser(req.files))
  packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(filePath, buffer)
    res.download(filePath)
  })
  dirCleaner(tempFilesDir)
});

module.exports = router;
