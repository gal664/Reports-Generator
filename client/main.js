let inputs = {
      text: {},
      file: {},
      reportTypes: {}
}
let tableauButton
let submitButton
let selectedReportType

document.addEventListener("DOMContentLoaded", () => {

      inputs.reportTypes.parent = document.querySelector("#selectReportType")
      inputs.reportTypes.types = inputs.reportTypes.parent.children
      tableauButton = document.querySelector("#tableauButton")

      for (let i = 0; i < inputs.reportTypes.types.length; i++) {

            let type = inputs.reportTypes.types[i]
            type.addEventListener("click", () => {
                  setTimeout(() => {
                        if (type.classList.contains("active")) selectedReportType = type.id
                        console.log(type.id)
                        if (selectedReportType != null) {
                              tableauButton.classList.remove("disabled")
                              switch (selectedReportType) {
                                    case "assessmentReport":
                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/331/views")
                                          break;
                                    case "practiceReport":
                                          tableauButton.setAttribute("href", "https://bi.timetoknow.co.il/#/workbooks/334/views")
                                          break;
                                    default:

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

      inputs.file.gradesBySubject = document.getElementById("grades_by_subject")
      inputs.file.gradesByQuestion = document.getElementById("grades_by_question")
      inputs.file.strugglingStudents = document.getElementById("struggling_students")
      submitButton = document.getElementById("submitButton")

      document.addEventListener("change", () => {
            if (selectedReportType != null
                  && inputs.file.gradesBySubject.files.length > 0
                  && inputs.file.gradesByQuestion.files.length > 0
                  && inputs.file.strugglingStudents.files.length > 0
                  && inputs.text.school.value != ""
                  && inputs.text.grade.value != ""
                  && inputs.text.classes.value != ""
                  && inputs.text.reportDate.value != ""
                  && inputs.text.assessmentName.value != "") submitButton.disabled = false
            else submitButton.disabled = true
      })

})