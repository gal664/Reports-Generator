const fs = require("fs");
const path = require("path");

module.exports = function (files) {
  
      // the array of file objects that will be returned
      let filesArray = []
    
      // runs through the objects that contains all of the files from multer
      Object.keys(files).forEach(key => {
        
        // refers to .csv files downloaded from tableau, specifically for student reports.
        // these files are comma-delimited, and some values contain commas themselves, those are inside double quotes
        if (files[key][0].fieldname == "student_data") {
          
          // the object that will contain the file's data
          let fileObject = {
            name: key,
            data: { students: [] }
          }
    
          // reading the file, encoding in UTF-8, since were using hebrew
          let decoded = fs.readFileSync(path.join(__dirname, "../", files[key][0].path), {encoding: "utf8"})
          
          // first of all, split the file in all of the newlines
          let decodedSplit = decoded.split("\r\n")
    
          // run a loop on the new array created above, which now contains strings, each string is a line delimited by ","
          for (let i = 1; i < decodedSplit.length - 1; i++) {
            
            // second, split the string only where there is a comma NOT inside a double quotes
            let splitRow = decodedSplit[i].match(/("[^"]*")|[^,]+/g)
    
            // start building the object. this only happens on the first loop
            if (!fileObject.data.assessmentTitle) fileObject.data.assessmentTitle = splitRow[0]
            if (!fileObject.data.grade) fileObject.data.grade = splitRow[2]
            if (!fileObject.data.schoolName) fileObject.data.schoolName = splitRow[3]
            if (!fileObject.data.rangeNumber) fileObject.data.rangeNumber = splitRow[5]
            
            // on each run of the loop, try to find the current student in the array of students
            let studentIndex = fileObject.data.students.find(student => student.name == splitRow[6])
    
            // if the student was not found AND there is a 7th value in the row, push the student to the array
            if (!studentIndex && splitRow[6]) {
              fileObject.data.students.push({
                name: splitRow[6],
                averageStudentScore: splitRow[8],
                studentStudyClassName: splitRow[4],
                subjects: [{ name: splitRow[1], verbalScore: splitRow[7] }]
              })
            
            // if the student was found AND there is a 7th value in the row,
            // find the subject in the student's subject array
            } else if (splitRow[6]) {
              let subjectIndex = studentIndex.subjects.find(subject => subject.name == splitRow[1])
              
              // if the subject was not found in the array, push the subject's values to the array
              if (!subjectIndex && splitRow[1])
                studentIndex.subjects.push({
                  name: splitRow[1],
                  verbalScore: splitRow[7]
                })
            }
          }
          
          // finally, push fileObject into the array
          filesArray.push(fileObject)
        
        // refers to .csv files downloaded from tableau, specifically for recommendations reports.
        // these files are comma-delimited, and some values contain commas themselves, those are inside double quotes.
        // it's important to clarify, these files are not at all different from student report files, but the object we build is slightly different
        } else if (files[key][0].fieldname == "recommendations_data") {
          
          // the object that will contain the file's data
          let fileObject = {
            name: key,
            data: { students: [] }
          }
          
          // reading the file, encoding in UTF-8, since were using hebrew
          let decoded = fs.readFileSync(path.join(__dirname, "../", files[key][0].path), {encoding: "utf8"})
          
          // first of all, split the file in all of the newlines
          let decodedSplit = decoded.split("\r\n")
    
          // run on the new array created above, which now contains strings, each string is a line delimited by ","
          for (let i = 1; i < decodedSplit.length - 1; i++) {
    
            // to split the string only where there is a comma NOT inside a double quotes
            let splitRow = decodedSplit[i].match(/("[^"]*")|[^,]+/g)
    
            // start building the object. this only happens on the first loop
            if (!fileObject.data.assessmentTitle) fileObject.data.assessmentTitle = splitRow[0]
            if (!fileObject.data.schoolName) fileObject.data.schoolName = splitRow[1]
            
            // on each run of the loop, try to find the current student in the array of students
            let studentIndex = fileObject.data.students.find(student => student.name == splitRow[5])
    
            // if the student was not found AND there is a 5th value in the row, push the student to the array
            if (!studentIndex && splitRow[5]) {
              fileObject.data.students.push({
                name: splitRow[5],
                studentStudyClassName: splitRow[2],
                Recommendations: [splitRow[3]]
              })
    
            // if the student was found AND there is a 5th value in the row,
            // find the recommendation in the student's recommendation array
            } else if (splitRow[5]) {
              let RecommendationsIndex = studentIndex.Recommendations.find(Recommendations => Recommendations == splitRow[3])
    
              // if the recommendation was not found in the array, push the recommendation's value to the array
              if (!RecommendationsIndex && splitRow[3])
                studentIndex.Recommendations.push(splitRow[3]);
            }
          }
    
          // finally, push fileObject into the array
          filesArray.push(fileObject);
        
        // refers to .csv files downloaded from tableau for any other type of report that is not recommendations or students.
        // these files are tab-delimited, so they are less problamatic to parse.
        } else {
          
          // the object that will contain the file's data
          let fileObject = {
            data: [],
            name: key
          }
          
          // reading the file, this time the encoding is UCS-2, which is fairly the only encoding that worked when parsing the file, since were using hebrew
          let decoded = fs.readFileSync(path.join(__dirname, "../", files[key][0].path), {encoding: "ucs2"})
          
          // first of all, split the file in all of the newlines
          let decodedSplit = decoded.split("\r\n");
    
          // run on the new array created above, which now contains strings, each string is a line delimited by "\t"
          for (let i = 0; i < decodedSplit.length; i++) {
            
            // split the string on all "\t" (tab) occurrences
            let splitRow = decodedSplit[i].split("\t");
            
            // if the splitRow's length is more than one item and the first item is not an empty string - push splitRow into the fileObject's data
            if (splitRow.length > 0 && splitRow[0] !== "")
              fileObject.data.push(splitRow)
          }
    
          // finally, push fileObject into the array 
          filesArray.push(fileObject);
        }
      })
    
      // return the files array after the parsing and organizing is done
      return filesArray
    }