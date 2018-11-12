const express = require("express")
const bodyParser = require("body-parser")
const path = require("path")
const app = express()
const reports = require("./reports")
const port = 9090

app.use(bodyParser.json())

app.use('/', express.static(path.join(__dirname, '../client')))

app.use("/reports", reports)

app.listen(port)

console.log(`server is listening at port ${port}`)