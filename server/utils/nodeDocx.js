const fs = require("fs")
const docx = require("docx")
const sheetLayoutConfig = require("../utils/sheetLayoutConfig")

const logoPath = sheetLayoutConfig.logoImagePath

module.exports = {

      studentReport(metadata, data) {

            let doc = new docx.Document()

            let assessmentTitle = data[0].data.assessmentTitle
            let grade = data[0].data.grade
            let schoolName = data[0].data.schoolName
            let students = data[0].data.students

            for (let i = 0; i < students.length; i++) {

                  let studentName = students[i].name
                  let averageScore = students[i].averageStudentScore
                  let studyclassName = students[i].studentStudyClassName
                  let subjects = students[i].subjects

                  doc.createImage(fs.readFileSync(logoPath))

                  addParagraphString(doc, `שלום ${studentName},`, "right", false)
                  addParagraphString(doc, `לפניך משוב על הישגיך במבחן "${assessmentTitle}"`, "right", false)
                  addParagraphString(doc, `ציונך במבחן: ${averageScore}`, "right", false)

                  // insert subjects table

                  for (let j = 0; j < subjects.length; j++) {

                        let subjectName = subjects[j].name
                        let verbalScore = subjects[j].verbalScore
                        let numeralScore = getNumeralScore(verbalScore)

                  }

                  // insert strings
                  addParagraphString(doc, "בהצלחה רבה,", "center", false)
                  addParagraphString(doc, "צוות עת הדעת", "center", false)
                  addParagraphString(doc, "עת הדעת | טלפון: 073-277-4800 | support-il@timetoknow.co.il | www.timetoknow.co.il", "center", true)
            }

            return doc
      }

}

function addParagraphString(doc, string, alignment, isPageBreak){
      
      let text = new docx.TextRun(string).size(24).font("calibri").rightToLeft()
      let paragraph

      if(alignment == "center"){

            if(isPageBreak){
                  paragraph = new docx.Paragraph().center().pageBreak()
            } else {
                  paragraph = new docx.Paragraph().center()
            }
      
      } else if(alignment == "right"){
            
            if(isPageBreak){
                  paragraph = new docx.Paragraph().right().pageBreak()
            } else {
                  paragraph = new docx.Paragraph().right()
            }
      }


      paragraph.addRun(text)
      doc.addParagraph(paragraph)
}

function getNumeralScore(score) {
      switch (score) {
            case "במידה מועטה מאוד": return "0-40"
            case "במידה מועטה": return "41-60"
            case "במידה חלקית": return "61-75"
            case "במידה רבה": return "76-85"
            case "במידה רבה מאוד": return "86-100"
      }
}