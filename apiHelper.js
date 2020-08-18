const ExcelJS = require('exceljs');

module.exports = {
    createWorkbook: function(res, result) {
        //console.log(result.data)
          //var worksheet = workbook.addWorksheet('some-worksheet')
          //worksheet.addRow(['foo', 'bar']).commit()
          addDataToWorkBook(res, result);
    }
};

function addDataToWorkBook(res, result) {
    var workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res })
    participantNumber = 1;
    result.data.forEach(participant  => {
        console.log(participant.id)
         var worksheet = workbook.addWorksheet("Participant " + participantNumber.toString());
         participantNumber++;
         worksheet.addRow(['foo', 'bar']).commit();

         worksheet.commit()
     });

    workbook.commit()
}