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

      // function to create an excel file with grade assessment data. should be used in /reports/gradeassessment request
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

      // function to create an excel file with practice data. should be used in /reports/practice request
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

                  // loops through the row cells
                  for (let j = 0; j < row.length; j++) {

                        // if the report is a grade assessment report, push an object to classMappingCalculations
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
                        
                        // insert the cell value from the data into the appropriate cell
                        sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))
                        
                        // if the cell is a number and is not a header
                        if (parseInt(row[j]) !== NaN && i > 0) {
                        
                              // if the value is under 58 - apply style redCellFill to the cell
                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                        
                              // if the value is between 58 and 74 - apply style orangeCellFill to the cell
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                        
                              // if the value is between 74 and 85 - apply style yellowCellFill to the cell
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                        
                              // if the value is over 85 - apply style greenCellFill to the cell
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }


                  }

            }
            
            // if the report is a grade assessment report - add to the counter in the matching object in classMappingCalculations
            if (reportType == "gradeAssessment") {
                  
                  // loop through the data rows
                  for (let i = 0; i < data.length; i++) {
                        
                        // easier to read
                        let row = data[i]

                        // loop through the first 3 cells of the row
                        for (let j = 0; j < 3; j++) {
                              
                              // easier to read
                              let cell = row[j]

                              // if the cell value is a number
                              if (parseInt(cell) !== NaN) {
                                    
                                    // if the value is between 0 and 58
                                    if (cell >= 0 && cell <= 58) {

                                          // loop through the objects in classMappingCalculations
                                          for (var x in classMappingCalculations) {
                                                
                                                // if the object name matches the value in row[0] (student's name)
                                                if (classMappingCalculations[x].name == row[0]) {
                                                      // add to the counter
                                                      classMappingCalculations[x].under58++
                                                      break;
                                                }
                                          }

                                    }
                                    // if the value is between 58 and 74
                                    else if (cell >= 59 && cell <= 74) {

                                          // loop through the objects in classMappingCalculations
                                          for (var x in classMappingCalculations) {
                                                
                                                // if the object name matches the value in row[0] (student's name)
                                                if (classMappingCalculations[x].name == row[0]) {
                                                
                                                      // add to the counter
                                                      classMappingCalculations[x].over58++
                                                      break;
                                                }
                                          }

                                    }
                                    // if the value is between 74 and 85
                                    else if (cell >= 74 && cell <= 85) {

                                          // loop through the objects in classMappingCalculations
                                          for (var x in classMappingCalculations) {
                                               
                                                // if the object name matches the value in row[0] (student's name)
                                                if (classMappingCalculations[x].name == row[0]) {
                                                     
                                                      // add to the counter
                                                      classMappingCalculations[x].over74++
                                                      break;
                                                }
                                          }

                                    }
                                    // if the value is between 85 and 100
                                    else if (cell >= 85 && cell <= 100) {

                                          // loop through the objects in classMappingCalculations
                                          for (var x in classMappingCalculations) {
                                               
                                                // if the object name matches the value in row[0] (student's name)
                                                if (classMappingCalculations[x].name == row[0]) {
                                                  
                                                      // add to the counter
                                                      classMappingCalculations[x].over85++
                                                      break;
                                                }
                                          }
                                    }
                              }
                        }
                  }
            }

            // if the report is a grade assessment report, change the values of the header's titles, A11, B11 to fit the report
            if (reportType == "gradeAssessment") {
                  
                  sheet.cell(startingRow, 1).string("כיתה")
                  sheet.cell(startingRow, 2).string("שם התלמיד")
                  sheet.cell(startingRow + 1, 1).string("ציון ממוצע")
                  sheet.cell(startingRow + 1, 2).string("ציון ממוצע")
                  
            // if the report is of any type other than grade assessment report, change the values of the header's titles
            } else {

                  sheet.cell(startingRow, 1).string("שם התלמיד")
                  sheet.cell(startingRow, 2).string("ציון סופי")

            }

      // inserts the data from grades_by_question file data into the sheet
      } else if (sheetType == "grades_by_question") {

            // loops through the data rows
            for (let i = 0; i < data.length; i++) {

                  // easier to read
                  let row = data[i]

                  // sets the height for the headers' row on the first run of the loop                  
                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  // loops through the row cells                  
                  for (let j = 0; j < row.length; j++) {

                        // if the current row is the first and the current cell is the third cell of the row or higher - change the title to be "שאלה - number"
                        if (i == 0 && j >= 2) sheet.cell(i + startingRow, j + 1).string(`שאלה ${j - 1}`).style(wb.createStyle(styles.reportData))
                        
                        // else just insert the cell value
                        else sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        // if the cell is a number and is not a header
                        if (parseInt(row[j]) !== NaN && i > 0) {
                        
                              // if the value is under 58 - apply style redCellFill to the cell
                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                        
                              // if the value is between 58 and 74 - apply style orangeCellFill to the cell
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                        
                              // if the value is between 74 and 85 - apply style yellowCellFill to the cell
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                        
                              // if the value is over 85 - apply style greenCellFill to the cell
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }

                  }

            }

            // if the report is a grade assessment report, change the values of the header's titles, A11, B11 to fit the report
            if (reportType == "gradeAssessment") {
                  
                  sheet.cell(startingRow, 1).string("כיתה")
                  sheet.cell(startingRow, 2).string("שם התלמיד")
                  sheet.cell(startingRow + 1, 1).string("ציון ממוצע")
                  sheet.cell(startingRow + 1, 2).string("ציון ממוצע")
                  
            // if the report is of any type other than grade assessment report, change the values of the header's titles
            } else {

                  sheet.cell(startingRow, 1).string("שם התלמיד")
                  sheet.cell(startingRow, 2).string("ציון סופי")

            }
      
      } else if (sheetType == "student_mapping") {

            // loops through the data rows
            for (let i = 0; i < data.length; i++) {

                  // sets the height for the headers' row on the first run of the loop                  
                  sheet.row(i + startingRow).setHeight(75)

                  // loops through the row cells                  
                  for (let j = 0; j < data[i].length; j++) {

                        // if both i & j are larger than 0 (not a name column or header row)
                        if (i > 0 && j > 0) {

                              // if cell value is not an empty string, insert "☆" in the cell and style it appropriately
                              if (data[i][j] != "") sheet.cell(j + startingRow, i + 1)
                              .string("☆")
                              .style(wb.createStyle(styles.reportData))
                              .style(wb.createStyle(styles.fontSize20pt))
                              
                              // if cell value is an empty string, insert "" in the cell and style it appropriately
                              if (data[i][j] == "") sheet.cell(j + startingRow, i + 1)
                                    .string("")
                                    .style(wb.createStyle(styles.reportData))

                        // if the i & j are not larger than 0 (student names or headers)
                        } else {
            
                              // insert the cell's data and style appropriately
                              sheet.cell(j + startingRow, i + 1)
                                    .string(data[i][j])
                                    .style(wb.createStyle(styles.reportData))

                        }
                  }
            }
            
            // change cell A10 to "שם התלמיד"
            sheet.cell(startingRow, 1).string("שם התלמיד")
      
      
      } else if (sheetType == "groups_by_subject") {

            // define variables for later use
            let longestRow
            
            // because the table should contain "lists" of students under each subject, the data should look a bit differently so it's easier to use
            let newData = []

            // loops through the data rows            
            for (let i = 0; i < data.length; i++) {

                  // loops through the row cells                  
                  for (let j = 1; j < data[i].length; j++) {

                        // easier to read
                        let cell = data[i][j]

                        // if the row is headers row
                        if (i == 0) {
                              
                              // create header object
                              let header = { name: cell, data: [] }
                              
                              // push header object to newData 
                              newData.push(header)
                        
                        // if the row is not a headers row
                        } else {
                              // if the cell velue is not an empty string - push it to the correct object in newData
                              if (cell != "") newData[j - 1].data.push(cell)
                              // reset the value of longestRow variable to match current row (currently longest)
                              longestRow = i
                        }
                  }
            }

            // sets the height for the headers' row on the first run of the loop                  
            sheet.row(startingRow).setHeight(100)

            // loop through the objects of newData
            for (let i = 0; i < newData.length; i++) {

                  // loop through the headers in each object
                  for (let j = 0; j < newData[i].data.length; j++) {

                        // if it's the first run of the loop, insert the header name to it's place
                        if (j == 0) {
                              sheet.cell(j + startingRow, i + 1)
                                    .string(newData[i].name)
                                    .style(wb.createStyle(styles.reportData))
                        }

                        // insert the student name to the appropriate cell and style it to have border only on left and right sides
                        sheet.cell(j + startingRow + 1, i + 1)
                        .string(newData[i].data[j])
                        .style(wb.createStyle(styles.reportDataNoBorderTopAndBottom))
                  }
                  
                  // style the last row that has data in it to have border on left, right and bottom sides
                  sheet.cell(startingRow + longestRow, i + 1)
                        .style(wb.createStyle(styles.reportDataNoBorderTop))
            }

            // go over the cells again except the last row and style them to have border only on left and right sides
            for (let i = 0; i < longestRow; i++) {

                  sheet.cell(startingRow + i, newData.length)
                        .style(wb.createStyle(styles.reportDataNoBorderTopAndBottom))

            }
      } else if (sheetType == "class_grades_by_subject") {

            // loop for 6 times in order to create the first table
            for (let i = 0; i < 6; i++) {
                  
                  // loop through classMappingCalculations' objects
                  for (let j = 0; j < classMappingCalculations.length; j++) {
                        
                        // define sum counters for each category
                        let sumOver85 = 0
                        let sumOver74 = 0
                        let sumOver58 = 0
                        let sumUnder58 = 0
                        let overallSum = 0
                        
                        // loop through classMappingCalculations' objects again and add each counter value to the matching sum counter
                        for (let x = 0; x < classMappingCalculations.length; x++) {
                              sumOver85 += classMappingCalculations[x].over85
                              sumOver74 += classMappingCalculations[x].over74
                              sumOver58 += classMappingCalculations[x].over58
                              sumUnder58 += classMappingCalculations[x].under58
                              
                              // define a total sum counter for each object of classMappingCalculations
                              classMappingCalculations[x].sum = classMappingCalculations[x].over85 + classMappingCalculations[x].over74 + classMappingCalculations[x].over58 + classMappingCalculations[x].under58
                        }

                        // add the sum of all of the object's counters to overallSum
                        overallSum += sumOver85 + sumOver74 + sumOver58 + sumUnder58

                        // if the row is headers row, insert appropriate values
                        if (i == 0) {
                              sheet.cell(i + startingRow, 1).string("טווח").style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 2).string(`סה"כ`).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).string("אחוזים").style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, j + 4).string(classMappingCalculations[j].name).style(wb.createStyle(styles.reportData))
                        
                        // insert the values of sumOver85 to the second row
                        } else if (i == 1) {
                              sheet.cell(i + startingRow, 1).string(">85")
                              .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.yellowCellFill))
                              sheet.cell(i + startingRow, 2).number(sumOver85).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumOver85 / overallSum)
                              .style(wb.createStyle(styles.reportData))
                              .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].over85).style(wb.createStyle(styles.reportData))
                        
                        // insert the values of sumOver74 to the third row
                        } else if (i == 2) {
                              sheet.cell(i + startingRow, 1).string("74-85")
                              .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.greenCellFill))
                              sheet.cell(i + startingRow, 2).number(sumOver74).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumOver74 / overallSum)
                              .style(wb.createStyle(styles.reportData))
                              .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].over74).style(wb.createStyle(styles.reportData))
                        
                        // insert the values of sumOver58 to the fourth row
                        } else if (i == 3) {
                              sheet.cell(i + startingRow, 1).string("58-73")
                              .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.orangeCellFill))
                              sheet.cell(i + startingRow, 2).number(sumOver58).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumOver58 / overallSum)
                              .style(wb.createStyle(styles.reportData))
                              .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].over58).style(wb.createStyle(styles.reportData))
                        
                        // insert the values of sumUnder58 to the fifth row
                        } else if (i == 4) {
                              sheet.cell(i + startingRow, 1).string("<58")
                              .style(wb.createStyle(styles.reportData)).style(wb.createStyle(styles.redCellFill))
                              sheet.cell(i + startingRow, 2).number(sumUnder58).style(wb.createStyle(styles.reportData))
                              sheet.cell(i + startingRow, 3).number(sumUnder58 / overallSum)
                              .style(wb.createStyle(styles.reportData))
                              .style(wb.createStyle(styles.percenatage))
                              sheet.cell(i + startingRow, j + 4).number(classMappingCalculations[j].under58).style(wb.createStyle(styles.reportData))
                        
                        // insert the values of overallSum to the sixth row
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

            // loop through the data rows, this time to create the second table on the sheet
            for (let i = 0; i < data.length; i++) {

                  // easier to read
                  let row = data[i]

                  // sets the height for the headers' row on the first run of the loop                  
                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  // loop through the row's first 2 cells
                  for (let j = 0; j < 2; j++) {
                        
                        // insert the cell value to the matching cell
                        sheet.cell(j + 18, i + 1).string(row[j]).style(wb.createStyle(styles.reportData))
                        
                        // if the cell is a number and is not a header
                        if (parseInt(row[j]) !== NaN && i > 0) {

                              // if the value is under 58 - apply style redCellFill to the cell
                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.redCellFill))

                              // if the value is between 58 and 74 - apply style orangeCellFill to the cell
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.orangeCellFill))

                              // if the value is between 74 and 85 - apply style yellowCellFill to the cell
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.yellowCellFill))

                              // if the value is over 85 - apply style greenCellFill to the cell
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(j + 18, i + 1).style(wb.createStyle(styles.greenCellFill))

                        }

                  }
            }

            // change cell A18 to "כיתה"
            sheet.cell(18, 1).string("כיתה")

            // reset startingRow to 22 (so the data starts 2 rows after the second table)
            startingRow = 22

            //loop through the data rows
            for (let i = 0; i < data.length; i++) {

                  // easier to read
                  let row = data[i]

                  // sets the height for the headers' row on row 10 on the first run of the loop                  
                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  // loop through the row cells
                  for (let j = 0; j < row.length; j++) {

                        // insert cell value to appropriate cell
                        sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        // if the cell is a number and is not a header
                        if (parseInt(row[j]) !== NaN && i > 0) {

                              // if the value is under 58 - apply style redCellFill to the cell
                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                        
                              // if the value is between 58 and 74 - apply style orangeCellFill to the cell
                              else if (row[j] >= 59 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                        
                              // if the value is between 74 and 85 - apply style yellowCellFill to the cell
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                        
                              // if the value is over 85 - apply style greenCellFill to the cell
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }

                  }

            }

            // change the values of the header's titles
            sheet.cell(startingRow, 1).string("כיתה")
            sheet.cell(startingRow, 2).string("ציון ממוצע")

      // this creates the last table on the third sheet of grade assessment report
      } else if (sheetType == "class_grades_by_question") {

            // set the starting row 2 lines after the third table
            startingRow = 2 + 22 + data.length

            // loop through the data rows
            for (let i = 0; i < data.length; i++) {

                  // easier to read
                  let row = data[i]

                  // sets the height for the headers' row on the first run of the loop
                  if (i == 0) sheet.row(startingRow).setHeight(100)

                  // loop through the row cells
                  for (let j = 0; j < row.length; j++) {

                        // if the current row is the first and the current cell is the third cell of the row or higher - change the title to be "שאלה - number"                        
                        if (i == 0 && j >= 2) sheet.cell(i + startingRow, j + 1).string(`שאלה ${j - 1}`).style(wb.createStyle(styles.reportData))

                        // else just insert the cell value
                        else sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))

                        // if the cell is a number and is not a header
                        if (parseInt(row[j]) !== NaN && i > 0) {

                              // if the value is under 58 - apply style redCellFill to the cell
                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                        
                              // if the value is between 58 and 74 - apply style orangeCellFill to the cell
                              else if (row[j] >= 58 && row[j] <= 74) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                        
                              // if the value is between 74 and 85 - apply style yellowCellFill to the cell
                              else if (row[j] >= 74 && row[j] <= 85) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                        
                              // if the value is over 85 - apply style greenCellFill to the cell
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))

                        }

                  }
            
            }
            // change the values of the header's titles
            sheet.cell(startingRow, 1).string("כיתה")
            sheet.cell(startingRow, 2).string("ציון ממוצע")

      }

}