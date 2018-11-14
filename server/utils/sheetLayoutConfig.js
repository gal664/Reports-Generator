let config = {
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
      heatMapIndexData: [
            { start: { row: 2, col: 8 }, end: { row: 2, col: 9 }, merged: true, content: "מקרא" },
            { start: { row: 3, col: 8 }, content: "צבע" },
            { start: { row: 3, col: 9 }, content: "טווח ציונים" },
            { start: { row: 4, col: 9 }, content: "85<" },
            { start: { row: 5, col: 9 }, content: "74-84" },
            { start: { row: 6, col: 9 }, content: "59-73" },
            { start: { row: 7, col: 9 }, content: "<58" }
      ],
      studentMappingIndexData: [
            { start: { row: 2, col: 8 }, end: { row: 2, col: 10 }, merged: true, content: "מקרא" },
            { start: { row: 3, col: 8 }, content: "סימון" },
            { start: { row: 3, col: 9 }, end: { row: 3, col: 10 }, merged: true, content: "משמעות" },
            { start: { row: 4, col: 8 }, content: "☆" },
            { start: { row: 4, col: 9 }, end: { row: 4, col: 10 }, merged: true, content: "צריך לבצע תרגול נוסף - ציון נמוך מ-57" },
      ],
      groupsBySubjectIndexData: [
            { start: { row: 2, col: 8 }, end: { row: 2, col: 10 }, merged: true, content: "מקרא" },
            { start: { row: 3, col: 8 }, end: { row: 3, col: 10 }, merged: true, content: "רשימת תלמידים לפי נושא שקיבלו ציון מתחת ל-57" },
      ]
}
module.exports = config