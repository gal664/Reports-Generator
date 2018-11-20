const styles = {
      whiteBorder: {
            border: {
                  left: { style: "thin", color: "white" },
                  right: { style: "thin", color: "white" },
                  top: { style: "thin", color: "white" },
                  bottom: { style: "thin", color: "white" }
            }
      },
      mediumBlackBorder: {
            border: {
                  left: { style: "medium", color: "black" },
                  right: { style: "medium", color: "black" },
                  top: { style: "medium", color: "black" },
                  bottom: { style: "medium", color: "black" }
            }
      },
      centerBold: {
            alignment: { vertical: "center", horizontal: "center" },
            font: { bold: true }
      },
      metadata: {
            alignment: { horizontal: "right" },
            font: { bold: true, underline: true }
      },
      reportData: {
            border: {
                  left: { style: "thin", color: "black" },
                  right: { style: "thin", color: "black" },
                  top: { style: "thin", color: "black" },
                  bottom: { style: "thin", color: "black" }
            },
            font: { bold: true },
            alignment: {
                  vertical: "center",
                  horizontal: "center",
                  wrapText: true
            }
      },
      reportDataNoBorderTop: {
            border: {
                  left: { style: "thin", color: "black" },
                  right: { style: "thin", color: "black" },
                  bottom: { style: "thin", color: "black" }
            },
            font: { bold: true },
            alignment: {
                  vertical: "center",
                  horizontal: "center",
                  wrapText: true
            }
      },
      reportDataNoBorderTopAndBottom: {
            border: {
                  left: { style: "thin", color: "black" },
                  right: { style: "thin", color: "black" }
            },
            font: { bold: true },
            alignment: {
                  vertical: "center",
                  horizontal: "center",
                  wrapText: true
            }
      },
      greenCellFill: {
            fill: {
                  type: "pattern",
                  patternType: "solid",
                  fgColor: "#92d050",
            },
            font: { color: "white" }
      },
      yellowCellFill: {
            fill: {
                  type: "pattern",
                  patternType: "solid",
                  fgColor: "#fffa00",
            }
      },
      orangeCellFill: {
            fill: {
                  type: "pattern",
                  patternType: "solid",
                  fgColor: "#ffbe00",
            }
      },
      redCellFill: {
            fill: {
                  type: "pattern",
                  patternType: "solid",
                  fgColor: "#ff0000",
            },
            font: { color: "white" }
      },
      fontSize20pt: {
            font: { size: 14, bold: false }
      },
      percenatage: { numberFormat: '#.00%; -#.00%; -' }
};

module.exports = styles;
