document.addEventListener("DOMContentLoaded", () => {
      let form = document.getElementById("report_upload_form")
      form.addEventListener("change", () => {
            let SCHOOL = document.getElementById("school").value
            let GRADE = document.getElementById("grade").value
            let CLASSES = document.getElementById("classes").value
            let REPORTDATE = document.getElementById("report_date").value
            let ASSESSMENTNAME = document.getElementById("assessment_name").value

            let actionStr = `/reports?assessmentname=${ASSESSMENTNAME}&school=${SCHOOL}&grade=${GRADE}&classes=${CLASSES}&reportdate=${REPORTDATE}`
            form.setAttribute("action", actionStr)
      })
})