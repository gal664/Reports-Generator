const csv = require('csvtojson')

module.exports = {
      // csvToJson(data){
      //       let rows = data.split("\n")
      //       let headers = rows[0].split("\t")
      //       let array = [{ headers: headers }]

      //       for (let i = 1; i < rows.length; ++i) {
      //             let cell = {}
      //             let l = rows[i].split("\t")

      //             for (let j = 0; j < headers.length; ++j) {
      //                   cell[headers[j]] = l[j]
      //             }
      //             array.push(cell)
      //       }

      //       return array
      // }
      
      csvToJson(path) {
            csv()
                  .fromFile(path)
                  .then((jsonObj) => {
                        return jsonObj
                  })
      }
}