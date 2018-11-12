var xl = require('excel4node')

module.exports = {
      assessmentReport(data) {
            // create the report excel workbook
            var wb = new xl.Workbook();
            
            //add spreadsheets
            var sheet1 = wb.addWorksheet('sheet1');
            var sheet2 = wb.addWorksheet('sheet2');
            var sheet3 = wb.addWorksheet('sheet3');

            // place the data
            console.log(data)
            sheet1.cell(1,1).string(data)

            // set styles

            // return finished report
            return(wb)
      }
}