const express = require("express")
const bodyParser = require("body-parser")
const path = require("path")
const app = express()
const reports = require("./index")

app.use(bodyParser.json())

app.use("/reports", reports)

app.use('/', express.static(path.join(__dirname, '../client')))

app.listen(9090)
