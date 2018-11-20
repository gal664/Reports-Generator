const fs = require("fs");
const path = require("path");

module.exports = function (dirPath) {

      // this timeout exist in order to wait 2 seconds before clearing all files in the directory
      // since there might be asyncronous actions that will result in files being saved
      setTimeout(() => {
            
            // reading the directory that is passed to the function
            fs.readdir(dirPath, (err, files) => {
                  // handle errors
                  if (err) throw err
                  
                  // run on all files in the directory
                  for (let file of files) {

                        // remove the files from the directory
                        fs.unlink(path.join(dirPath, file), err => {
                              
                              //handle errors
                              if (err) throw err
                        })
                  }
            })
      }, 2000)
}