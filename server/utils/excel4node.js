const xl = require('excel4node')
const wbMetadata = {
      logoImagePath: "./utils/t2k_logo.png",
      schoolData: [
            { startCell: "A2", content: "עיר:" },
            { startCell: "A3", content: "בית ספר:" },
            { startCell: "A4", content: "שכבה:" },
            { startCell: "A5", content: "כיתות:" },
            { startCell: "A6", content: "תקופת הדוח:" }
      ],
      indexData: [
            { startCell: "D1", endCell: "E1", merged: true, content: "מקרא" },
            { startCell: "E2", content: "טווח ציונים" },
            { startCell: "D2", content: "צבע" },
            { startCell: "E3", content: "85<" },
            { startCell: "E4", content: "74-84" },
            { startCell: "E5", content: "59-73" },
            { startCell: "E6", content: "<58" },
      ]
}

module.exports = {

      assessmentReport(metadata, data) {
            let wb = new xl.Workbook()

            // for(item of data){
            for (let i = 0; i < data.length; i++) {
                  switch (data[i].name) {
                        case "grades_by_subject":
                              let grades_by_subject_sheet = createSheetLayout(wb, metadata, "ציוני תרגול")
                              insertReportData(grades_by_subject_sheet, data[i].data, "grades_by_subject")
                              break;
                        case "struggling_students":
                              let struggling_students_sheet = createSheetLayout(wb, metadata, "תרגול נוסף לפי נושאים")
                              break;
                        case "grades_by_question":
                              let grades_by_question_sheet = createSheetLayout(wb, metadata, "תלמידים מתקשים לפי נושא")
                              break;
                  }
            }

            return wb
      }
}

function createSheetLayout(wb, metadata, sheetName) {

      let sheetOptions = {
            sheetView: {
                  rightToLeft: true,
            },
            sheetFormat: {
                  defaultColWidth: 18,
            },
      }

      let sheet = wb.addWorksheet(sheetName, sheetOptions)
      let sheetBasicStyle = wb.createStyle({
            border: {
                  left: { style: "thin", color: "white" },
                  right: { style: "thin", color: "white" },
                  top: { style: "thin", color: "white" },
                  bottom: { style: "thin", color: "white" },
            }
      })

      sheet.cell(1, 1, 500, 100).style(sheetBasicStyle)
      sheet.addImage({
            path: wbMetadata.logoImagePath,
            type: 'picture',
            position: {
                  type: 'oneCellAnchor',
                  from: {
                        col: 9,
                        colOff: 0,
                        row: 1,
                        rowOff: 0,
                  }
            },
      })

      // sheet metadata
      let metadataStyle = wb.createStyle({
            alignment: { horizontal: "right" },
            font: { bold: true, underline: true }
      })

      for (item of wbMetadata.schoolData) {
            contentToCell(sheet, item)
            applyStyleToCell(sheet, item, metadataStyle)
      }

      insertMetadataValues(sheet, metadata)

      // sheet index data
      let indexStyle = wb.createStyle({
            alignment: { horizontal: "center" },
            font: { bold: true }
      })

      for (item of wbMetadata.indexData) {
            contentToCell(sheet, item)
            applyStyleToCell(sheet, item, indexStyle)
      }
      let indexBorderStyle = wb.createStyle({
            border: {
                  left: { style: "medium", color: "black" },
                  right: { style: "medium", color: "black" },
                  top: { style: "medium", color: "black" },
                  bottom: { style: "medium", color: "black" },
            }
      })

      sheet.cell(1, 4, 6, 5).style(indexBorderStyle)
      let greenStyle = wb.createStyle({ fill: { type: "pattern", patternType: "solid", fgColor: "#92d050" } })
      let yellowStyle = wb.createStyle({ fill: { type: "pattern", patternType: "solid", fgColor: "#fffa00" } })
      let orangeStyle = wb.createStyle({ fill: { type: "pattern", patternType: "solid", fgColor: "#ffbe00" } })
      let redStyle = wb.createStyle({ fill: { type: "pattern", patternType: "solid", fgColor: "#ff0000" } })
      sheet.cell(3, 4).style(greenStyle)
      sheet.cell(4, 4).style(yellowStyle)
      sheet.cell(5, 4).style(orangeStyle)
      sheet.cell(6, 4).style(redStyle)

      return sheet
}

function contentToCell(sheet, cellOptions) {    // { startCell: "D1", endCell: "E1", merged: true, content: "מקרא" }
      if (!cellOptions.endCell) {
            let cellLocation = xl.getExcelRowCol(cellOptions.startCell)
            sheet.cell(cellLocation.row, cellLocation.col).string(cellOptions.content)
      } else {
            let cellLocation = xl.getExcelRowCol(cellOptions.startCell)
            let endCellLocation = xl.getExcelRowCol(cellOptions.endCell)
            sheet.cell(cellLocation.row, cellLocation.col, endCellLocation.row, endCellLocation.col, cellOptions.merged).string(cellOptions.content)
      }
}

function applyStyleToCell(sheet, cellOptions, style) {
      let cellLocation = xl.getExcelRowCol(cellOptions.startCell)
      sheet.cell(cellLocation.row, cellLocation.col).style(style)
}

function insertMetadataValues(sheet, metadata) {
      contentToCell(sheet, { startCell: "B2", content: metadata.city })
      contentToCell(sheet, { startCell: "B3", content: metadata.school })
      contentToCell(sheet, { startCell: "B4", content: metadata.grade })
      contentToCell(sheet, { startCell: "B5", content: metadata.classes })
      contentToCell(sheet, { startCell: "B6", content: metadata.reportperiod })
}

function insertReportData(sheet, data, sheetType) {
      switch (sheetType) {
            case "grades_by_subject":
            for(let i = 0; i < data.length; i++){
                  let row = data[i]
                        for(let j = 0; j < row.length; j++){
                              sheet.cell(i + 9, j + 1).string(row[j])
                        }
                  }
                  contentToCell(sheet, { startCell: "A9", content: "שם התלמיד" })
                  contentToCell(sheet, { startCell: "B9", content: "ציון סופי" })
                  break;
      }
}