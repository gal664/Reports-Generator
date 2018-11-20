let inputs = {
      text: {},
      file: {},
      reportTypes: {}
}

let tableauButton
let submitButton
let selectedReportType
let fileInputsContainer

document.addEventListener("DOMContentLoaded", () => {

      inputs.reportTypes.parent = document.querySelector("#selectReportType")
      inputs.reportTypes.types = inputs.reportTypes.parent.children

      tableauButton = document.querySelector("#tableauButton")

      fileInputsContainer = document.querySelector("#fileInputs")
      inputs.file.grades_by_subject = createFileInputElement("grades_by_subject","ציוני תלמיד לפי מבחן")
      inputs.file.grades_by_question = createFileInputElement("grades_by_question","ציוני תלמיד לפי שאלה")
      inputs.file.struggling_students = createFileInputElement("struggling_students","תלמידים מתקשים לפי נושאים")
      inputs.file.class_grades_by_subject = createFileInputElement("class_grades_by_subject","ציוני כיתה לפי מבחן")
      inputs.file.class_grades_by_question = createFileInputElement("class_grades_by_question","ציוני כיתה לפי שאלה")
      inputs.file.student_data = createFileInputElement("student_data","נתוני תלמידים")
      inputs.file.recommendations_data = createFileInputElement("recommendations_data","המלצות לתלמידים")

      for (var inputElement in inputs.file) {
            
            if (inputs.file.hasOwnProperty(inputElement)) {
                  
                  let element = inputs.file[inputElement]
                  
                  element.addEventListener("change", (event) => {
                              
                        let input = element.children[0]
                        let label = element.children[1]
                        if (element.firstChild.files.length > 0) {
                              label.innerHTML = input.files[0].name
                              label.style.color = "#00d326"
                        }
                        
                  })
            }
      }

      for (let i = 0; i < inputs.reportTypes.types.length; i++) {

            let type = inputs.reportTypes.types[i]
            type.addEventListener("click", () => {
                  
                  setTimeout(() => {
                        
                        selectedReportType = type.id
                        console.log(selectedReportType)

                        if (selectedReportType != null) {

                              tableauButton.classList.remove("disabled")

                              switch (selectedReportType) {
                                    case "assessment":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.firstChild.firstChild.value = ""
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/331/views")

                                          fileInputsContainer.appendChild(inputs.file.grades_by_subject)
                                          fileInputsContainer.appendChild(inputs.file.grades_by_question)
                                          fileInputsContainer.appendChild(inputs.file.struggling_students)

                                          break;

                                    case "practice":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.firstChild.firstChild.value = ""
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/334/views")

                                          fileInputsContainer.appendChild(inputs.file.grades_by_subject)
                                          fileInputsContainer.appendChild(inputs.file.struggling_students)

                                          break;

                                    case "gradeAssessment":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.firstChild.firstChild.value = ""
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/353/views")

                                          fileInputsContainer.appendChild(inputs.file.grades_by_subject)
                                          fileInputsContainer.appendChild(inputs.file.grades_by_question)
                                          fileInputsContainer.appendChild(inputs.file.class_grades_by_subject)
                                          fileInputsContainer.appendChild(inputs.file.class_grades_by_question)

                                          break;

                                    case "student":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.firstChild.firstChild.value = ""
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/333/views")

                                          fileInputsContainer.appendChild(inputs.file.student_data)

                                          break;

                                    case "recommendations":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.firstChild.firstChild.value = ""
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/330/views")

                                          fileInputsContainer.appendChild(inputs.file.recommendations_data)

                                          break;
                              }
                        }
                  }, 1);
            })
      }

      inputs.text.school = document.getElementById("school")
      inputs.text.grade = document.getElementById("grade")
      inputs.text.classes = document.getElementById("classes")
      inputs.text.reportDate = document.getElementById("report_date")
      inputs.text.assessmentName = document.getElementById("assessment_name")

      let form = document.getElementById("report_upload_form")

      form.addEventListener("change", () => {
            let actionStr = `/reports/${selectedReportType}?school=${inputs.text.school.value}&grade=${inputs.text.grade.value}&classes=${inputs.text.classes.value}&reportdate=${inputs.text.reportDate.value}&assessmentname=${inputs.text.assessmentName.value}`
            form.setAttribute("action", actionStr)
      })

      submitButton = document.getElementById("submitButton")

      let eventsArray = ["click", "change", "mouseup", "mousedown", "mousemove"]

      eventsArray.forEach(eventType => {

            document.addEventListener(eventType, () => {

                  if (selectedReportType == "assessment"
                        && inputs.file.grades_by_subject.firstElementChild.files.length > 0
                        && inputs.file.grades_by_question.firstElementChild.files.length > 0
                        && inputs.file.struggling_students.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else if (selectedReportType == "practice"
                        && inputs.file.grades_by_subject.firstElementChild.files.length > 0
                        && inputs.file.struggling_students.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else if (selectedReportType == "gradeAssessment"
                        && inputs.file.grades_by_subject.firstElementChild.files.length > 0
                        && inputs.file.grades_by_question.firstElementChild.files.length > 0
                        && inputs.file.class_grades_by_subject.firstElementChild.files.length > 0
                        && inputs.file.class_grades_by_question.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else if (selectedReportType == "student"
                        && inputs.file.student_data.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else if (selectedReportType == "recommendations"
                        && inputs.file.recommendations_data.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else submitButton.disabled = true

            })

      })

})

function createFileInputElement(fileType, text) {

      let container = document.createElement("div")
      container.className = "custom-file mb-3"

      let input = document.createElement("input")
      input.className = "custom-file-input"
      input.id = fileType
      input.name = fileType
      input.setAttribute("type", "file")
      input.setAttribute("accept", ".csv")
      container.appendChild(input)

      let label = document.createElement("label")
      label.className = "custom-file-label"
      label.setAttribute("for", fileType)
      label.setAttribute("originallabeltext", text)
      label.innerHTML = text
      container.appendChild(label)

      return container
}

function createTextInputElement(valueName, value) {

      let container = document.createElement("div")
      container.className = "form-group mb-3"

      let input = document.createElement("input")
      input.className = "form-control"
      input.id = valueName
      input.placeholder = value
      input.setAttribute("type", "text")
      container.appendChild(input)
      
      return container
}