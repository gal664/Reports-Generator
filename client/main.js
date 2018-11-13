document.addEventListener("DOMContentLoaded", () => {
      let form = document.getElementById("report_upload_form")
      form.addEventListener("change", () => {
            let CITY = document.getElementById("city").value
            let SCHOOL = document.getElementById("school").value
            let GRADE = document.getElementById("grade").value
            let CLASSES = document.getElementById("classes").value
            let REPORTPERIOD = document.getElementById("report_period").value

            let actionStr = `/reports?city=${CITY}&school=${SCHOOL}&grade=${GRADE}&classes=${CLASSES}&reportperiod=${REPORTPERIOD}`
            console.log(actionStr)

            form.setAttribute("action", actionStr)
      })
})