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

let uploadInputs = upload.fields([
  { name: "grades_by_subject", maxCount: 1 },
  { name: "struggling_students", maxCount: 1 },
  { name: "grades_by_question", maxCount: 1 },
  { name: "class_grades_by_subject", maxCount: 1 },
  { name: "class_grades_by_question", maxCount: 1 },
  { name: "student_data", maxCount: 1 },
  { name: "recommendations_data", maxCount: 1 },
]);

router.post("/assessment", uploadInputs, (req, res) => {
  if (!req.files) {
    reject("no files uploaded");
    res.status(500).send("error uploading the files");
  }

  let wb = xlsx.assessmentReport(req.query, requestFilesToArrays(req.files));
  wb.write(
    `${req.query.classes} - ${req.query.assessmentname} - דוח מבחן.xlsx`,
    res
  );

  deleteAllTempFiles();
});

router.post("/practice", uploadInputs, (req, res) => {
  if (!req.files) {
    reject("no files uploaded");
    res.status(500).send("error uploading the files");
  }

  let wb = xlsx.practiceReport(req.query, requestFilesToArrays(req.files));
  wb.write(
    `${req.query.classes} - ${req.query.assessmentname} - דוח תרגול.xlsx`,
    res
  );

  deleteAllTempFiles();
});

router.post("/gradeAssessment", uploadInputs, (req, res) => {
  if (!req.files) {
    reject("no files uploaded");
    res.status(500).send("error uploading the files");
  }

  let wb = xlsx.gradeAssessmentReport(
    req.query,
    requestFilesToArrays(req.files)
  );
  wb.write(
    `${req.query.classes} - ${req.query.assessmentname} - דוח מבחן שכבתי.xlsx`,
    res
  );

  deleteAllTempFiles();
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
    requestFilesToArrays(req.files)
  );

  packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(filePath, buffer);
    res.download(filePath);
  });

  deleteAllTempFiles();
});

router.post("/recommendations", uploadInputs, (req, res) => {
  const packer = new docx.Packer()
  const fileName = `${req.query.classes} - ${req.query.assessmentname} - דוח המלצות.docx`
  const filePath = path.join(tempFilesDir, fileName)
  if (!req.files) {
    reject("no files uploaded")
    res.status(500).send("error uploading the files")
  }

  let doc = docxFunctions.recommendationsReport(req.query, requestFilesToArrays(req.files))
  packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(filePath, buffer)
    res.download(filePath)
  })
  deleteAllTempFiles()
});

function requestFilesToArrays(files) {
  let filesArray = [];

  Object.keys(files).forEach(key => {
    if (files[key][0].fieldname == "student_data") {
      let fileObject = { name: key };
      let decoded = fs.readFileSync(path.join(__dirname, files[key][0].path), {
        encoding: "utf8"
      });
      let decodedSplit = decoded.split("\r\n");
      let data = {
        students: []
      };

      for (let i = 1; i < decodedSplit.length - 1; i++) {
        
        // to split the string only where there is a comma NOT inside a double quotes
        let splitRow = decodedSplit[i].match(/("[^"]*")|[^,]+/g);

        if (!data.assessmentTitle) data.assessmentTitle = splitRow[0];
        if (!data.grade) data.grade = splitRow[2];
        if (!data.schoolName) data.schoolName = splitRow[3];
        if (!data.rangeNumber) data.rangeNumber = splitRow[5];

        let studentIndex = data.students.find(
          student => student.name == splitRow[6]
        );

        if (!studentIndex && splitRow[6]) {
          data.students.push({
            name: splitRow[6],
            averageStudentScore: splitRow[8],
            studentStudyClassName: splitRow[4],
            subjects: [{ name: splitRow[1], verbalScore: splitRow[7] }]
          });
        } else if (splitRow[6]) {
          let subjectIndex = studentIndex.subjects.find(
            subject => subject.name == splitRow[1]
          );

          if (!subjectIndex && splitRow[1])
            studentIndex.subjects.push({
              name: splitRow[1],
              verbalScore: splitRow[7]
            });
        }
      }

      fileObject.data = data;
      filesArray.push(fileObject);
    } else if (files[key][0].fieldname == "recommendations_data") {
      let fileObject = { name: key };
      let decoded = fs.readFileSync(path.join(__dirname, files[key][0].path), {
        encoding: "utf8"
      });
      let decodedSplit = decoded.split("\r\n");
      
      let data = {
        students: []
      };

      for (let i = 1; i < decodedSplit.length - 1; i++) {

        // to split the string only where there is a comma NOT inside a double quotes
        let splitRow = decodedSplit[i].match(/("[^"]*")|[^,]+/g);

        if (!data.assessmentTitle) data.assessmentTitle = splitRow[0];
        if (!data.schoolName) data.schoolName = splitRow[1];
        
        let studentIndex = data.students.find(
          student => student.name == splitRow[5]
        );

        if (!studentIndex && splitRow[5]) {
          data.students.push({
            name: splitRow[5],
            studentStudyClassName: splitRow[2],
            Recommendations: [splitRow[3]]
          });
        } else if (splitRow[5]) {
          let RecommendationsIndex = studentIndex.Recommendations.find(
            Recommendations => Recommendations == splitRow[3]
          );

          if (!RecommendationsIndex && splitRow[3])
            studentIndex.Recommendations.push(splitRow[3]);
        }
      }
      
      fileObject.data = data;
      filesArray.push(fileObject);
    } else {
      let fileObject = { data: [], name: key };
      let decoded = fs.readFileSync(path.join(__dirname, files[key][0].path), {
        encoding: "ucs2"
      });
      let decodedSplit = decoded.split("\r\n");

      for (let i = 0; i < decodedSplit.length; i++) {
        let splitRow = decodedSplit[i].split("\t");
        if (splitRow.length > 0 && splitRow[0] !== "")
          fileObject.data.push(splitRow);
      }

      filesArray.push(fileObject);
    }
  });

  return filesArray;
}

function deleteAllTempFiles() {
  fs.readdir(tempFilesDir, (err, files) => {
    if (err) throw err;
    for (const file of files) {
      fs.unlink(path.join(tempFilesDir, file), err => {
        if (err) throw err;
      });
    }
  });
}

module.exports = router;
