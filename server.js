'use strict'
const express = require('express')
const app = express()
const port = 5000

const axios = require('axios');
var helper = require("./apiHelper");

 app.get('/', (req, res) => {
    
    axios.get('http://experiment-database.herokuapp.com/results').then(result => {
        res.writeHead(200, {
            'Content-Disposition': 'attachment; filename="exp.xlsx"',
            'Transfer-Encoding': 'chunked',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          })
          helper.createWorkbook(res, result);
    }).catch(err => {
        console.log(err)
    })
})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
});