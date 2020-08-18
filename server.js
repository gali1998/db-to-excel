'use strict'
const express = require('express')
const app = express()
const port = 5000
const ExcelJS = require('exceljs');

 app.get('/', (req, res) => {
    res.writeHead(200, {
        'Content-Disposition': 'attachment; filename="file.xlsx"',
        'Transfer-Encoding': 'chunked',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      })
      var workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res })
      var worksheet = workbook.addWorksheet('some-worksheet')
      worksheet.addRow(['foo', 'bar']).commit()
      worksheet.commit()
      workbook.commit()
})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
});