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
      inputs.file.grades_by_subject = createFileInputElement("grades_by_subject")
      inputs.file.grades_by_question = createFileInputElement("grades_by_question")
      inputs.file.struggling_students = createFileInputElement("struggling_students")

      for (let i = 0; i < inputs.reportTypes.types.length; i++) {

            let type = inputs.reportTypes.types[i]
            type.addEventListener("click", () => {

                  setTimeout(() => {

                        if (type.classList.contains("active")) selectedReportType = type.id

                        if (selectedReportType != null) {

                              tableauButton.classList.remove("disabled")

                              switch (selectedReportType) {
                                    case "assessmentReport":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/331/views")

                                          fileInputsContainer.appendChild(inputs.file.grades_by_subject)
                                          fileInputsContainer.appendChild(inputs.file.grades_by_question)
                                          fileInputsContainer.appendChild(inputs.file.struggling_students)

                                          break;

                                    case "practiceReport":

                                          while (fileInputsContainer.childNodes.length != 0) {
                                                fileInputsContainer.removeChild(fileInputsContainer.firstChild)
                                          }

                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/334/views")

                                          fileInputsContainer.appendChild(inputs.file.grades_by_subject)
                                          fileInputsContainer.appendChild(inputs.file.struggling_students)

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

                  if (selectedReportType == "assessmentReport"
                        && inputs.file.grades_by_subject.firstElementChild.files.length > 0
                        && inputs.file.grades_by_question.firstElementChild.files.length > 0
                        && inputs.file.struggling_students.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else if (selectedReportType == "practiceReport"
                        && inputs.file.grades_by_subject.firstElementChild.files.length > 0
                        && inputs.file.struggling_students.firstElementChild.files.length > 0
                        && inputs.text.school.value != ""
                        && inputs.text.grade.value != ""
                        && inputs.text.classes.value != ""
                        && inputs.text.reportDate.value != ""
                        && inputs.text.assessmentName.value != "") submitButton.disabled = false
                  else submitButton.disabled = true

            })

      })

})

function createFileInputElement(fileType) {

      let container = document.createElement("div")
      container.className = "custom-file mb-3"

      let input = document.createElement("input")
      input.className = "custom-file-input"
      input.id = fileType
      input.name = fileType
      input.setAttribute("type", "file")
      container.appendChild(input)

      let label = document.createElement("label")
      label.className = "custom-file-label"
      label.setAttribute("for", fileType)
      label.innerHTML = fileType
      container.appendChild(label)

      return container
}