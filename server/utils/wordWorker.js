const fs = require("fs")
const docx = require("docx")
const sheetLayoutConfig = require("./sheetLayoutConfig")
const figures = require('figures')

const logoPath = sheetLayoutConfig.logoImagePath

module.exports = {

      // function to create a word file with student data. should be used in /reports/student request      
      studentReport(data) {

            // create a new instance of a docx document with custom margins in options
            let doc = new docx.Document(undefined, { top: 500, right: 750, bottom: 500, left: 750 })

            // define variables for all of the general data values
            let assessmentTitle = data[0].data.assessmentTitle
            let grade = data[0].data.grade
            let schoolName = data[0].data.schoolName
            let students = data[0].data.students
            
            // loop through students array
            for (let i = 0; i < students.length; i++) {
                  
                  // define variables for all of the student's data values
                  let studentName = students[i].name
                  let averageScore = students[i].averageStudentScore
                  let studyclassName = students[i].studentStudyClassName
                  let subjects = students[i].subjects

                  // define logo size variables
                  let LogoWidth = 605
                  let LogoHeight = 250

                  // insert logo to the document
                  doc.createImage(fs.readFileSync(logoPath), LogoWidth/3, LogoHeight/3)
                  
                  // align the first page with all the rest since later on we add another line that will make all of them drop a line
                  if(i == 0) addParagraphString(doc, "", "center", false)

                  // add text to document using the variables defined above
                  addParagraphString(doc, `שלום ${studentName},`, "right", false)
                  addParagraphString(doc, `לפניך משוב על הישגיך במבחן ${assessmentTitle}`, "right", false)
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, `ציונך במבחן: ${averageScore}`, "right", false)
                  addParagraphString(doc, "", "center", false)
                  
                  // *********************IMPORTANT*********************
                  // since were using hebrew the table starts on the opposite direction.
                  // this means that the loops are reversed (kind of),
                  // when examining the loops - take this into account
                  // ***************************************************
                  
                  // add the grades table
                  const table = doc.createTable(subjects.length + 1, 6)
                  
                  // set the table width
                  table.setWidth(docx.WidthType.PERCENTAGE, '100%');
                  
                  // define the table's headers
                  let tableHeaders = [
                        "במידה מועטה מאוד",
                        "במידה מועטה",
                        "במידה חלקית",
                        "במידה רבה",
                        "במידה רבה מאוד"
                  ]

                  // loop through the student's subjects
                  for(let x = 0; x < subjects.length; x++){
                       
                        
                        // define variables for all of the subject's data values
                        let subjectName = subjects[x].name
                        let verbalScore = subjects[x].verbalScore
                        let numeralScore = getNumeralScore(verbalScore)

                        // insert '✔︎' sign in the correct place that matches the subject's score
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

                        // loop 6 times inside the subject's loop
                        for(let j = 0; j < 6; j++){
                              // if it's the first subject - add the headers
                              if(x == 0) addTableCell(table, x, j, tableHeaders[j])
                              // else if it's the first (last) cell - insert the subject's name
                              if(j == 5) addTableCell(table, x + 1, j, subjectName)
                              
                        }

                  }

                  // insert strings after the table
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "בהצלחה רבה,", "center", false)
                  addParagraphString(doc, "צוות עת הדעת", "center", false)
                  
                  // this is the blank pageBreak mentioned earlier,
                  // this is a walkaround since adding a pageBreak with text makes the text skip to the next page as well
                  addParagraphString(doc, "", "center", true)
            }

            // add a string variable for the footer string
            let footerString = "עת הדעת | טלפון: 073-277-4800 | support-il@timetoknow.co.il | www.timetoknow.co.il"

            // add the string with some styles to a new instance of docx textRun
            let footerTextRun = new docx.TextRun(footerString).size(24).font("calibri").rightToLeft()
            
            // add a footer for the document with the textRun defined earlier and center it
            doc.Footer.createParagraph(footerTextRun).center()

            // return the ready document
            return doc
      },
      recommendationsReport(data) {

            // create a new instance of a docx document with custom margins in options
            let doc = new docx.Document(undefined, { top: 500, right: 750, bottom: 500, left: 750 })

            // define variables for all of the general data values
            let assessmentTitle = data[0].data.assessmentTitle
            let schoolName = data[0].data.schoolName
            let students = data[0].data.students

            // loop through students array
            for (let i = 0; i < students.length; i++) {
                  
                  // define variables for all of the student's data values
                  let studentName = students[i].name
                  let studyclassName = students[i].studentStudyClassName
                  let Recommendations = students[i].Recommendations

                  // define logo size variables
                  let LogoWidth = 605
                  let LogoHeight = 250

                  // insert logo to the document
                  let logo = doc.createImage(fs.readFileSync(logoPath), LogoWidth/3, LogoHeight/3)
                  
                  // align the first page with all the rest since later on we add another line that will make all of them drop a line
                  if(i == 0) addParagraphString(doc, "", "right", false)

                  // add text to document using the variables defined above
                  addParagraphString(doc, `שלום ${studentName},`, "right", false)
                  addParagraphString(doc, "", "right", false)
                  addParagraphString(doc, "לפניך שמות תרגולים שאנו ממליצים לך לתרגל כדי לשלוט טוב יותר בנושאים הנלמדים.", "right", false)
                  addParagraphString(doc, `התרגולים נמצאים באתר עת הדעת בקורס "${assessmentTitle}".`, "right", false)
                  addParagraphString(doc, "כל תרגול כולל בתוכו מספר תרגילים וניתן לעבוד עליהם בהמשכים.", "right", false)
                  addParagraphString(doc, "להלן רשימת התרגולים המומלצים לך לתרגול:", "right", false)
                  addParagraphString(doc, "", "center", false)
                  
                  // loop through the student's recommendations
                  for(let x = 0; x < Recommendations.length; x++){

                        // define bullets for all of the recommendations.
                        // it's done using a string since the normal bullets are flipped and there was no visible solution
                        // for flipping them sides.
                        addParagraphString(doc, `     ●   ${Recommendations[x]}`, "right", false)
                  }

                  // insert strings after the bullet list
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "אם נתקלת בבעיה טכנית באפשרותך לפנות למוקד התמיכה שלנו בפרטים המופיעים מטה.", "right", false)
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "", "center", false)
                  addParagraphString(doc, "מאחלים לך הצלחה רבה,", "center", false)
                  addParagraphString(doc, "צוות עת הדעת", "center", false)

                  // this is the blank pageBreak mentioned earlier,
                  // this is a walkaround since adding a pageBreak with text makes the text skip to the next page as well
                  addParagraphString(doc, "", "center", true)
            }

            // add a string variable for the footer string
            let footerString = "עת הדעת | טלפון: 073-277-4800 | support-il@timetoknow.co.il | www.timetoknow.co.il"

            // add the string with some styles to a new instance of docx textRun
            let footerTextRun = new docx.TextRun(footerString).size(24).font("calibri").rightToLeft()
            
            // add a footer for the document with the textRun defined earlier and center it
            doc.Footer.createParagraph(footerTextRun).center()

            // return ready doc
            return doc
      }

}

// adds a cell to a table
function addTableCell(table, row, col, string){
      
      // creates an instance of docx TextRun with the given string, adds some style
      let text = new docx.TextRun(string).size(24).font("calibri").bold().rightToLeft()
      
      // creates an instance of docx Paragraph with some style
      let paragraph = new docx.Paragraph().center().spacing({before:250, after:250})
      
      // appends the textRun to the paragraph
      paragraph.addRun(text)
      
      // inserts the paragraph to the given location in the table
      table.getCell(row, col).addContent(paragraph)

}

// adds a paragraph to a doc
function addParagraphString(doc, string, alignment, isPageBreak){
      
      // creates an instance of docx TextRun with the given string, adds some style
      let text = new docx.TextRun(string).size(24).font("calibri").rightToLeft()
      
      let paragraph

      // for aligning the text to the center of the page
      if(alignment == "center"){

            // for creating a pageBreak
            if(isPageBreak){

                  // creates an instance of docx Paragraph with adds some style and pageBreak
                  paragraph = new docx.Paragraph().center().pageBreak()

            } else {
                  
                  // creates an instance of docx Paragraph with adds some style
                  paragraph = new docx.Paragraph().center()

            }
            
      // for aligning the text to the right side of the page      
      } else if(alignment == "right"){
            
            // for creating a pageBreak
            if(isPageBreak){

                  // creates an instance of docx Paragraph with adds some style and pageBreak
                  paragraph = new docx.Paragraph().right().pageBreak()

            } else {

                  // creates an instance of docx Paragraph with adds some style
                  paragraph = new docx.Paragraph().right()

            }

      }

      // appends the textRun to the paragraph
      paragraph.addRun(text)

      // inserts the paragraph to the document
      doc.addParagraph(paragraph)
}

// converts the verbal score to numeral score, returns the correct one
function getNumeralScore(score) {
      
      switch (score) {

            case "במידה מועטה מאוד": return "0-40"
            case "במידה מועטה": return "41-60"
            case "במידה חלקית": return "61-75"
            case "במידה רבה": return "76-85"
            case "במידה רבה מאוד": return "86-100"
      
      }

}