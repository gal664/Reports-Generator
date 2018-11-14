const xl = require('excel4node')
const wbMetadata = {
      wbOptions: {
            defaultFont: { size: 10 },
      },
      logoImagePath: "./utils/t2k_logo.png",
      schoolData: [
            { start: { row: 2, col: 4 }, content: "בית ספר:" },
            { start: { row: 3, col: 4 }, content: "שכבה:" },
            { start: { row: 4, col: 4 }, content: "כיתות:" },
            { start: { row: 5, col: 4 }, content: "תאריך המבחן:" },
            { start: { row: 6, col: 4 }, content: "שם המבחן:" }
      ],
      indexData: [
            { start: { row: 2, col: 8 }, end: { row: 2, col: 9 }, merged: true, content: "מקרא" },
            { start: { row: 3, col: 8 }, content: "צבע" },
            { start: { row: 3, col: 9 }, content: "טווח ציונים" },
            { start: { row: 4, col: 9 }, content: "85<" },
            { start: { row: 5, col: 9 }, content: "74-84" },
            { start: { row: 6, col: 9 }, content: "59-73" },
            { start: { row: 7, col: 9 }, content: "<58" }
      ]
}
const styles = require("./excelStyles")

module.exports = {

      assessmentReport(metadata, data) {

            let wb = new xl.Workbook(wbMetadata.wbOptions)

            for (let i = 0; i < data.length; i++) {
                  switch (data[i].name) {
                        case "grades_by_subject":
                              let sheet1 = createSheetLayout(wb, metadata, "מיפוי מיומנויות")
                              insertReportData(wb, sheet1, data[i].data, data[i].name)
                              break;
                        case "grades_by_question":
                              let sheet2 = createSheetLayout(wb, metadata, "ציונים לפי שאלה")
                              insertReportData(wb, sheet2, data[i].data, data[i].name)
                              break;
                        case "struggling_students":
                              let sheet3 = createSheetLayout(wb, metadata, "מיפוי לתלמיד")
                              insertReportData(wb, sheet3, data[i].data, data[i].name)

                              let sheet4 = createSheetLayout(wb, metadata, "קבוצות לפי נושא")
                              // insertReportData(wb, sheet4, data[i].data, data[i].name)
                              break;
                  }
            }

            return wb
      }
}

function createSheetLayout(wb, metadata, sheetName) {

      let sheetOptions = {
            sheetView: { rightToLeft: true },
            sheetFormat: { defaultColWidth: 18, defaultRowHeight: 18 }
      }

      let sheet = wb.addWorksheet(sheetName, sheetOptions)
      sheet.cell(1, 1, 500, 100).style(wb.createStyle(styles.whiteBorder))

      sheet.addImage({
            path: wbMetadata.logoImagePath,
            type: 'picture',
            position: { type: 'oneCellAnchor', from: { col: 1, colOff: 0, row: 1, rowOff: 0 } },
      })

      for (item of wbMetadata.schoolData) {
            sheet.cell(item.start.row, item.start.col)
                  .string(item.content)
                  .style(wb.createStyle(styles.metadata))
      }

      sheet.cell(2, 5).string(metadata.school)
      sheet.cell(3, 5).string(metadata.grade)
      sheet.cell(4, 5).string(metadata.classes)
      sheet.cell(5, 5).string(metadata.reportdate)
      sheet.cell(6, 5).string(metadata.assessmentname)

      for (item of wbMetadata.indexData) {
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

      return sheet
}

function insertReportData(wb, sheet, data, sheetType) {
      let startingRow = 10

      if (sheetType == "grades_by_subject") {
            for (let i = 0; i < data.length; i++) {
                  let row = data[i]
                  if (i == 0) sheet.row(startingRow).setHeight(100)
                  for (let j = 0; j < row.length; j++) {
                        sheet.cell(i + startingRow, j + 1).string(row[j]).style(wb.createStyle(styles.reportData))
                        if (parseInt(row[j]) !== NaN && i > 0) {
                              if (row[j] >= 0 && row[j] <= 58) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.redCellFill))
                              else if (row[j] >= 59 && row[j] <= 73) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 84) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))
                        }
                  }
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
                              else if (row[j] >= 59 && row[j] <= 73) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.orangeCellFill))
                              else if (row[j] >= 74 && row[j] <= 84) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.yellowCellFill))
                              else if (row[j] >= 85 && row[j] <= 100) sheet.cell(i + startingRow, j + 1).style(wb.createStyle(styles.greenCellFill))
                        }
                  }
            }
      } else if (sheetType == "struggling_students") {
            for (let i = 0; i < data.length; i++) {
                  sheet.row(i + startingRow).setHeight(75)

                  for (let j = 0; j < data[i].length; j++) {
                        if(i > 0 && j > 0){
                              if(data[i][j] != "") sheet.cell(j + startingRow, i + 1).string("☆").style(wb.createStyle(styles.reportData))
                              if(data[i][j] == "") sheet.cell(j + startingRow, i + 1).string("").style(wb.createStyle(styles.reportData))
                        } else {
                              sheet.cell(j + startingRow, i + 1).string(data[i][j]).style(wb.createStyle(styles.reportData))
                        }
                  }
            }
      }
      sheet.cell(startingRow, 1).string("שם התלמיד")
      sheet.cell(startingRow, 2).string("ציון סופי")
}