const xl = require('excel4node')
const styles = require("./excelStyles")
const wbMetadata = require("./sheetLayoutConfig")

module.exports = {
      // function to create an excel file with assessment data. should be used in /reports/assessment request
      assessmentReport(metadata, data) {
            
            //create a new workbook instance
            let wb = new xl.Workbook(wbMetadata.wbOptions)

            // loop through the objects made from the .csv files
            for (let i = 0; i < data.length; i++) {

                  switch (data[i].name) {

                        // creates the first sheet from "grade_by_subject" file
                        case "grades_by_subject":
                              
                              // create sheet in the workbook
                              let sheet1 = createSheetLayout(wb, metadata, "מיפוי מיומנויות", data[i].name)
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet1, data[i].data, data[i].name)
                              break;
                              
                        // creates the second sheet from "grades_by_question" file
                        case "grades_by_question":
                              
                              // create sheet in the workbook
                              let sheet2 = createSheetLayout(wb, metadata, "ציונים לפי שאלה", data[i].name)
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet2, data[i].data, data[i].name)
                              break;
                        
                        // creates the third and fourth sheets from "struggling_students" file
                        case "struggling_students":
                              
                              // create sheet in the workbook
                              let sheet3 = createSheetLayout(wb, metadata, "מיפוי לתלמיד", "student_mapping")
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet3, data[i].data, "student_mapping")
                              
                              // create sheet in the workbook
                              let sheet4 = createSheetLayout(wb, metadata, "קבוצות לפי נושא", "groups_by_subject")
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet4, data[i].data, "groups_by_subject")
                              break;
                  }
            }

            // return workbook when all sheets were created
            return wb
      },

      // function to create an excel file with assessment data. should be used in /reports/assessment request
      gradeAssessmentReport(metadata, data) {

            //create a new workbook instance
            let wb = new xl.Workbook(wbMetadata.wbOptions)

            // array to contain the re-ordered data for the third sheet
            let thirdSheetData = []

            // array to contain the calculations needed for the class mapping in the third sheet
            let classMappingCalculations = []

            // loop through the objects made from the .csv files
            for (let i = 0; i < data.length; i++) {

                  switch (data[i].name) {

                        // creates the first sheet from "grade_by_subject" file
                        case "grades_by_subject":
                              
                              // create sheet in the workbook
                              let sheet1 = createSheetLayout(wb, metadata, "מיפוי מיומנויות", data[i].name)
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet1, data[i].data, data[i].name, "gradeAssessment", classMappingCalculations)
                              break;

                        // creates the second sheet from "grades_by_question" file
                        case "grades_by_question":
                              
                              // create sheet in the workbook
                              let sheet2 = createSheetLayout(wb, metadata, "ציונים לפי שאלה", data[i].name)
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet2, data[i].data, data[i].name, "gradeAssessment")
                              break;

                        // insert the data from "class_grades_by_subject" and "class_grades_by_question" to array defined above
                        case "class_grades_by_subject":
                        case "class_grades_by_question":
                              thirdSheetData.push(data[i])
                              break;
                  }
            }
            
            // create the third sheet in the workbook
            let sheet3 = createSheetLayout(wb, metadata, "מיפוי שכבתי", "gradeMapping")
            
            // loop through the third sheet's data array
            for (let i = 0; i < thirdSheetData.length; i++) {

                  switch (thirdSheetData[i].name) {

                        case "class_grades_by_subject":
                        
                              // insert the data from "class_grades_by_subject" file to the third sheet
                              insertReportData(wb, sheet3, thirdSheetData[i].data, thirdSheetData[i].name, "gradeAssessment", classMappingCalculations)
                              break;
                              
                        case "class_grades_by_question":

                              // insert the data from "class_grades_by_question" file to the third sheet
                              insertReportData(wb, sheet3, thirdSheetData[i].data, thirdSheetData[i].name, "gradeAssessment", classMappingCalculations)
                              break;
                  }
            }

            // return workbook when all sheets were created
            return wb
      },

      // function to create an excel file with assessment data. should be used in /reports/assessment request
      practiceReport(metadata, data) {

            //create a new workbook instance
            let wb = new xl.Workbook(wbMetadata.wbOptions)

            // loop through the objects made from the .csv files
            for (let i = 0; i < data.length; i++) {

                  switch (data[i].name) {

                        // creates the first sheet from "grade_by_subject" file
                        case "grades_by_subject":
                              
                              // create sheet in the workbook
                              let sheet1 = createSheetLayout(wb, metadata, "מיפוי מיומנויות", data[i].name)
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet1, data[i].data, data[i].name)
                              break;

                        // creates the second and fourth sheets from "struggling_students" file
                        case "struggling_students":
                              
                              // create sheet in the workbook
                              let sheet2 = createSheetLayout(wb, metadata, "מיפוי לתלמיד", "student_mapping")
                              
                              // insert the data into the sheet
                              insertReportData(wb, sheet2, data[i].data, "student_mapping")
                              break;

                  }
            }

            // return workbook when all sheets were created
            return wb
      }
}

// creates the sheet without the data in it, sort of a template maker
function createSheetLayout(wb, metadata, sheetName, sheetType) {

      // define the options for the sheet
      let sheetOptions = {
            
            // define view to be right to left since were using hebrew
            sheetView: { rightToLeft: true },
            
            // define default column width and row height
            sheetFormat: { defaultColWidth: 18, defaultRowHeight: 18 }
      }

      // add a new sheet to the workbook using the options defined above and the name passed to the function
      let sheet = wb.addWorksheet(sheetName, sheetOptions)
      
      // this is a workaround, could not find a way to color ALL the border in all the cells white.
      // it would take too long to apply white thin border to all the cells,
      // so here we only apply it to cells 1-500 in rows 1-500
      sheet.cell(1, 1, 500, 100).style(wb.createStyle(styles.whiteBorder))

      // add the logo image to cell A1
      sheet.addImage({
            path: wbMetadata.logoImagePath,
            type: 'picture',
            position: { type: 'oneCellAnchor', from: { col: 1, colOff: 0, row: 1, rowOff: 0 } },
      })

      // insert metadata to cells defined in sheetLayoutConfig.schoolData (these are the metadata headers)
      for (item of wbMetadata.schoolData) {
            sheet.cell(item.start.row, item.start.col)
                  .string(item.content)
                  .style(wb.createStyle(styles.metadata))
      }

      // insert the metadata values to the appropriate cells
      sheet.cell(2, 5).string(metadata.school)
      sheet.cell(3, 5).string(metadata.grade)
      sheet.cell(4, 5).string(metadata.classes)
      sheet.cell(5, 5).string(metadata.reportdate)
      sheet.cell(6, 5).string(metadata.assessmentname)

      // create index cells for heatmap sheets
      if (sheetType == "gradeMapping" || sheetType == "grades_by_subject" || sheetType == "grades_by_question") {
            for (item of wbMetadata.heatMapIndexData) {
                  if (!item.merged) {
                        sheet.cell(item.start.row, item.start.col)
                              .string(item.content)
                              .style(wb.createStyle(styles.centerBold))
                  } else {
                        sheet.cell(item.start.row, item.start.col, item.end.row, item.end.col, item.merged)
                              .string(item.content)
                              .style(wb.createStyle(styles.centerBold))
                  }
            }

            sheet.cell(2, 8, 7, 9).style(wb.createStyle(styles.mediumBlackBorder))
            sheet.cell(4, 8).style(wb.createStyle(styles.greenCellFill))
            sheet.cell(5, 8).style(wb.createStyle(styles.yellowCellFill))
            sheet.cell(6, 8).style(wb.createStyle(styles.orangeCellFill))
            sheet.cell(7, 8).style(wb.createStyle(styles.redCellFill))
      
      // create index cells for student mapping sheet      
      } else if (sheetType == "student_mapping") {
            for (item of wbMetadata.studentMappingIndexData) {
                  if (!item.merged) {
                        sheet.cell(item.start.row, item.start.col)
                              .string(item.content)
                              .style(wb.createStyle(styles.centerBold))
                  } else {
                        sheet.cell(item.start.row, item.start.col, item.end.row, item.end.col, item.merged)
                              .string(item.content)
                              .style(wb.createStyle(styles.centerBold))
                  }
            }

            sheet.cell(4, 8).style(wb.createStyle(styles.fontSize20pt))
            sheet.cell(2, 8, 4, 10).style(wb.createStyle(styles.mediumBlackBorder))

      // create index cells for geoups by subjects sheet      
      } else if (sheetType == "groups_by_subject") {
            for (item of wbMetadata.groupsBySubjectIndexData) {
                  sheet.cell(item.start.row, item.start.col, item.end.row, item.end.col, item.merged)
                        .string(item.content)
                        .style(wb.createStyle(styles.centerBold))
            }

            sheet.cell(2, 8, 3, 10).style(wb.createStyle(styles.mediumBlackBorder))
      }

      // return the ready sheet
      return sheet
}
// receives a ready sheet with all the styling done and inserts the report data into the appropriate place in the sheet
function insertReportData(wb, sheet, data, sheetType, reportType, classMappingCalculations) {
      
      // defines a row for the data to start from.
      // changing this will cause the data to start from a different row
      let startingRow = 10

      // this array is not being used actively
      // it collects all the headers for the report to make sure they dont repeat
      let headersArray = []

      // inserts the data from grades_by_subject file data into the sheet
      if (sheetType == "grades_by_subject") {

            // loops through the data rows
            for (let i = 0; i < data.length; i++) {
                  
                  // easier to read that way
                  let row = data[i]
                  
                  // sets the height for the headers' row on the first run of the loop
                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  // loops through the row's cells
                  for (let j = 0; j < row.length; j++) {

                        
                        if (j == 0 && i > 1 && reportType == "gradeAssessment") {
                              if (headersArray.indexOf(row[j]) == -1) {
                                    headersArray.push(row[j])
                                    classMappingCalculations.push({
                                          name: row[j],
                                          over85: 0,
                                          over74: 0,
                                          over58: 0,
                                          under58: 0
                                    })
                              }
                        }

                        sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        if (parseInt(row[j]) !== NaN && i > 0) {

                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }


                  }

            }

            if (reportType == "gradeAssessment") {
                  for (let i = 0; i < data.length; i++) {
                        let row = data[i]

                        for (let j = 0; j < 3; j++) {
                              let cell = row[j]

                              if (parseInt(cell) !== NaN) {

                                    if (cell >= 0 && cell <= 58) {


                                          for (var x in classMappingCalculations) {
                                                if (classMappingCalculations[x].name == row[0]) {
                                                      classMappingCalculations[x].under58++
                                                      break;
                                                }
                                          }

                                    }

                                    else if (cell >= 59 && cell <= 74) {
                                          sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))

                                          for (var x in classMappingCalculations) {
                                                if (classMappingCalculations[x].name == row[0]) {
                                                      classMappingCalculations[x].over58++
                                                      break;
                                                }
                                          }

                                    }

                                    else if (cell >= 74 && cell <= 85) {
                                          sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))

                                          for (var x in classMappingCalculations) {
                                                if (classMappingCalculations[x].name == row[0]) {
                                                      classMappingCalculations[x].over74++
                                                      break;
                                                }
                                          }

                                    }

                                    else if (cell >= 85 && cell <= 100) {
                                          sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                                          for (var x in classMappingCalculations) {
                                                if (classMappingCalculations[x].name == row[0]) {
                                                      classMappingCalculations[x].over85++
                                                      break;
                                                }
                                          }
                                    }
                              }
                        }
                  }
            }

            if (reportType == "gradeAssessment") {

                  sheet.cell(startingRow, 1).string("כיתה")
                  sheet.cell(startingRow, 2).string("שם התלמיד")
                  sheet.cell(startingRow + 1, 1).string("ציון ממוצע")
                  sheet.cell(startingRow + 1, 2).string("ציון ממוצע")

            } else {

                  sheet.cell(startingRow, 1).string("שם התלמיד")
                  sheet.cell(startingRow, 2).string("ציון סופי")

            }

      } else if (sheetType == "grades_by_question") {

            for (let i = 0; i < data.length; i++) {

                  let row = data[i]

                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  for (let j = 0; j < row.length; j++) {

                        if (i == 0 && j >= 2) sheet.cell(i + startingRow, j + 1).string(`שאלה ${j - 1}`).style(wb.createStyle(styles.reportData))
                        else sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        if (parseInt(row[j]) !== NaN && i > 0) {

                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }

                  }

            }

            if (reportType == "gradeAssessment") {

                  sheet.cell(startingRow, 1).string("כיתה")
                  sheet.cell(startingRow, 2).string("שם התלמיד")
                  sheet.cell(startingRow + 1, 1).string("ציון ממוצע")
                  sheet.cell(startingRow + 1, 2).string("ציון ממוצע")

            } else {

                  sheet.cell(startingRow, 1).string("שם התלמיד")
                  sheet.cell(startingRow, 2).string("ציון סופי")

            }

      } else if (sheetType == "student_mapping") {
            for (let i = 0; i < data.length; i++) {

                  sheet.row(i + startingRow).setHeight(75)

                  for (let j = 0; j < data[i].length; j++) {

                        if (i > 0 && j > 0) {

                              if (data[i][j] != "") sheet.cell(j + startingRow, i + 1)
                                    .string("☆")
                                    .style(wb.createStyle(styles.reportData))
                                    .style(wb.createStyle(styles.fontSize20pt))

                              if (data[i][j] == "") sheet.cell(j + startingRow, i + 1)
                                    .string("")
                                    .style(wb.createStyle(styles.reportData))

                        } else {

                              sheet.cell(j + startingRow, i + 1)
                                    .string(data[i][j])
                                    .style(wb.createStyle(styles.reportData))

                        }
                  }
            }

            sheet.cell(startingRow, 1).string("שם התלמיד")

      } else if (sheetType == "groups_by_subject") {
            let longestRow

            let newData = []

            for (let i = 0; i < data.length; i++) {

                  for (let j = 1; j < data[i].length; j++) {

                        let cell = data[i][j]

                        if (i == 0) {
                              let header = { name: cell, data: [] }
                              newData.push(header)
                        } else {
                              if (cell != "") newData[j - 1].data.push(cell)
                              longestRow = i
                        }
                  }
            }

            sheet.row(startingRow).setHeight(100)

            for (let i = 0; i < newData.length; i++) {

                  for (let j = 0; j < newData[i].data.length; j++) {

                        if (j == 0) {
                              sheet.cell(j + startingRow, i + 1)
                                    .string(newData[i].name)
                                    .style(wb.createStyle(styles.reportData))
                        }

                        sheet.cell(j + startingRow + 1, i + 1)
                              .string(newData[i].data[j])
                              .style(wb.createStyle(styles.reportDataNoBorderTopAndBottom))
                  }

                  sheet.cell(startingRow + longestRow, i + 1)
                        .style(wb.createStyle(styles.reportDataNoBorderTop))
            }

            for (let i = 0; i < longestRow; i++) {

                  sheet.cell(startingRow + i, newData.length)
                        .style(wb.createStyle(styles.reportDataNoBorderTopAndBottom))

            }
      } else if (sheetType == "class_grades_by_subject") {

            for (let i = 0; i < 6; i++) {

                  for (let j = 0; j < classMappingCalculations.length; j++) {

                        let sumOver85 = 0
                        let sumOver74 = 0
                        let sumOver58 = 0
                        let sumUnder58 = 0
                        let overallSum = 0

                        for (let i = 0; i < classMappingCalculations.length; i++) {
                              sumOver85 += classMappingCalculations[i].over85
                              sumOver74 += classMappingCalculations[i].over74
                              sumOver58 += classMappingCalculations[i].over58
                              sumUnder58 += classMappingCalculations[i].under58
                              classMappingCalculations[i].sum = classMappingCalculations[i].over85 + classMappingCalculations[i].over74 + classMappingCalculations[i].over58 + classMappingCalculations[i].under58
                        }

                        overallSum += sumOver85 + sumOver74 + sumOver58 + sumUnder58

                        if (i == 0) {
                              sheet.cell(i + startingRow, 1).string("טווח").style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 2).string(`סה"כ`).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).string("אחוזים").style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, j + 4).string(classMappingCalculations[j].name).style(wb.createStyle(styles.reportData))
                        } else if (i == 1) {
                              sheet.cell(i + startingRow, 1).string(">85")
                                    .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.yellowCellFill))
                              sheet.cell(i + startingRow, 2).number(sumOver85).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumOver85 / overallSum)
                                    .style(wb.createStyle(styles.reportData))
                                    .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].over85).style(wb.createStyle(styles.reportData))
                        } else if (i == 2) {
                              sheet.cell(i + startingRow, 1).string("74-85")
                                    .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.greenCellFill))
                              sheet.cell(i + startingRow, 2).number(sumOver74).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumOver74 / overallSum)
                                    .style(wb.createStyle(styles.reportData))
                                    .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].over74).style(wb.createStyle(styles.reportData))
                        } else if (i == 3) {
                              sheet.cell(i + startingRow, 1).string("58-73")
                                    .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.orangeCellFill))
                              sheet.cell(i + startingRow, 2).number(sumOver58).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumOver58 / overallSum)
                                    .style(wb.createStyle(styles.reportData))
                                    .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].over58).style(wb.createStyle(styles.reportData))
                        } else if (i == 4) {
                              sheet.cell(i + startingRow, 1).string("<58")
                                    .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.redCellFill))
                              sheet.cell(i + startingRow, 2).number(sumUnder58).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumUnder58 / overallSum)
                                    .style(wb.createStyle(styles.reportData))
                                    .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].under58).style(wb.createStyle(styles.reportData))
                        } else if (i == 5) {
                              sheet.cell(i + startingRow, 1).string(`סה"כ`).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 2).number(overallSum).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(overallSum / overallSum)
                                    .style(wb.createStyle(styles.reportData))
                                    .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].sum).style(wb.createStyle(styles.reportData))
                        }

                  }

            }

            for (let i = 0; i < data.length; i++) {

                  let row = data[i]

                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  for (let j = 0; j < 2; j++) {

                        sheet.cell(j + 18, i + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        if (parseInt(row[j]) !== NaN && i > 0) {

                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.redCellFill))
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.greenCellFill))

                        }
                  }
            }


            sheet.cell(18, 1).string("כיתה")

            startingRow = 22

            for (let i = 0; i < data.length; i++) {

                  let row = data[i]

                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  for (let j = 0; j < row.length; j++) {

                        sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        if (parseInt(row[j]) !== NaN && i > 0) {

                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }
                  }
            }

            sheet.cell(startingRow, 1).string("כיתה")
            sheet.cell(startingRow, 2).string("ציון ממוצע")

      } else if (sheetType == "class_grades_by_question") {

            startingRow = 2 + 22 + data.length

            for (let i = 0; i < data.length; i++) {

                  let row = data[i]

                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  for (let j = 0; j < row.length; j++) {

                        if (i == 0 && j >= 2) sheet.cell(i + startingRow, j + 1).string(`שאלה ${j - 1}`).style(wb.createStyle(styles.reportData))
                        else sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        if (parseInt(row[j]) !== NaN && i > 0) {

                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                              else if (row[j] >= 58 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }
                  }
            }

            sheet.cell(startingRow, 1).string("כיתה")
            sheet.cell(startingRow, 2).string("ציון ממוצע")

      }

}