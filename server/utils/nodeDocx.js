const fs = require("fs")
const docx = require("docx")
const sheetLayoutConfig = require("../utils/sheetLayoutConfig")
const figures = require('figures');

const logoPath = sheetLayoutConfig.logoImagePath

module.exports = {

      studentReport(metadata, data) {

            let doc = new docx.Document(undefined, {
                  top: 500,
                  right: 750,
                  bottom: 500,
                  left: 750,
              })

            let assessmentTitle = data[0].data.assessmentTitle
            let grade = data[0].data.grade
            let schoolName = data[0].data.schoolName
            let students = data[0].data.students

            for (let i = 0; i < students.length; i++) {
                  
                  let studentName = students[i].name
                  let averageScore = students[i].averageStudentScore
                  let studyclassName = students[i].studentStudyClassName
                  let subjects = students[i].subjects

                  let LogoWidth = 605
                  let LogoHeight = 250
                  let logo = doc.createImage(fs.readFileSync(logoPath), LogoWidth/3, LogoHeight/3)
                  
                  if(i == 0) addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, `שלום ${studentName},`, "right", false)
                  addParagraphString(doc, `לפניך משוב על הישגיך במבחן ${assessmentTitle}`, "right", false)
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, `ציונך במבחן: ${averageScore}`, "right", false)
                  addParagraphString(doc, "", "center", false)
                  
                  const table = doc.createTable(subjects.length + 1, 6)
                  table.setWidth(docx.WidthType.PERCENTAGE, '100%');
                  
                  let tableHeaders = [
                        "במידה מועטה מאוד",
                        "במידה מועטה",
                        "במידה חלקית",
                        "במידה רבה",
                        "במידה רבה מאוד"
                  ]

                  
                  for(let x = 0; x < subjects.length; x++){
                       
                        let subjectName = subjects[x].name
                        let verbalScore = subjects[x].verbalScore
                        let numeralScore = getNumeralScore(verbalScore)

                        switch (verbalScore) {
                              case "במידה מועטה מאוד":
                                    addTableCell(table, x + 1, 0, figures('✔︎'))
                                    break;
                              case "במידה מועטה":
                                    addTableCell(table, x + 1, 1, figures('✔︎'))
                                    break;
                              case "במידה חלקית":
                                    addTableCell(table, x + 1, 2, figures('✔︎'))
                                    break;
                              case "במידה רבה":
                                    addTableCell(table, x + 1, 3, figures('✔︎'))
                                    break;
                              case "במידה רבה מאוד":
                                    addTableCell(table, x + 1, 4, figures('✔︎'))
                                    break;
                        }

                        for(let j = 0; j < 6; j++){

                              if(x == 0) addTableCell(table, x, j, tableHeaders[j])
                              if(j == 5) addTableCell(table, x + 1, j, subjectName)
                              
                        }

                  }

                  // insert strings
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "בהצלחה רבה,", "center", false)
                  addParagraphString(doc, "צוות עת הדעת", "center", false)
                  addParagraphString(doc, "", "center", true)
            }

            let footerString = "עת הדעת | טלפון: 073-277-4800 | support-il@timetoknow.co.il | www.timetoknow.co.il"
            let footerTextRun = new docx.TextRun(footerString).size(24).font("calibri").rightToLeft()
            doc.Footer.createParagraph(footerTextRun).center()

            return doc
      },
      recommendationsReport(metadata, data) {

            // console.log(data[0].data.students)
            // {
            //       name: 'שחר מאיה',
            //       studentStudyClassName: 'אנגלית ה1',
            //       Recommendations: ['Image Description  - Family and School']
            // }
            let doc = new docx.Document(undefined, {
                  top: 500,
                  right: 750,
                  bottom: 500,
                  left: 750,
              })

            let assessmentTitle = data[0].data.assessmentTitle
            let schoolName = data[0].data.schoolName
            let students = data[0].data.students

            for (let i = 0; i < students.length; i++) {
                  
                  let studentName = students[i].name
                  let studyclassName = students[i].studentStudyClassName
                  let Recommendations = students[i].Recommendations

                  let LogoWidth = 605
                  let LogoHeight = 250
                  let logo = doc.createImage(fs.readFileSync(logoPath), LogoWidth/3, LogoHeight/3)
                  
                  if(i == 0) addParagraphString(doc, "", "right", false)
                  addParagraphString(doc, `שלום ${studentName},`, "right", false)
                  addParagraphString(doc, "", "right", false)
                  addParagraphString(doc, "לפניך שמות תרגולים שאנו ממליצים לך לתרגל כדי לשלוט טוב יותר בנושאים הנלמדים.", "right", false)
                  addParagraphString(doc, `התרגולים נמצאים באתר עת הדעת בקורס "${assessmentTitle}".`, "right", false)
                  addParagraphString(doc, "כל תרגול כולל בתוכו מספר תרגילים וניתן לעבוד עליהם בהמשכים.", "right", false)
                  addParagraphString(doc, "להלן רשימת התרגולים המומלצים לך לתרגול:", "right", false)
                  addParagraphString(doc, "", "center", false)
                  
                  
                  for(let x = 0; x < Recommendations.length; x++){
                        
                        addParagraphString(doc, `     ●   ${Recommendations[x]}`, "right", false)

                  }

                  // insert strings
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "אם נתקלת בבעיה טכנית באפשרותך לפנות למוקד התמיכה שלנו בפרטים המופיעים מטה.", "right", false)
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "מאחלים לך הצלחה רבה,", "center", false)
                  addParagraphString(doc, "צוות עת הדעת", "center", false)
                  addParagraphString(doc, "", "center", true)
            }

            let footerString = "עת הדעת | טלפון: 073-277-4800 | support-il@timetoknow.co.il | www.timetoknow.co.il"
            let footerTextRun = new docx.TextRun(footerString).size(24).font("calibri").rightToLeft()
            doc.Footer.createParagraph(footerTextRun).center()

            return doc
      }

}

function addTableCell(table, row, col, string){
      
      let text = new docx.TextRun(string).size(24).font("calibri").bold().rightToLeft()
      let paragraph = new docx.Paragraph().center().spacing({before:250, after:250})
      paragraph.addRun(text)
      table.getCell(row, col).addContent(paragraph)

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