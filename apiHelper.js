const ExcelJS = require('exceljs');
var _ = require('lodash');
const { min } = require('lodash');
var moment = require('moment'); // require
const { time } = require('console');

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
        let headers = ["טקסט", "זמן מתחילת ניסוי", "קטגוריה"]

        // create worksheet for participant
         let v = [{state: 'normal', rightToLeft: true, topLeftCell: 'G1'}]
         var worksheet = workbook.addWorksheet("Participant " + participantNumber.toString(), {views: v});
         participantNumber++;

         // add id line on top
         worksheet.addRow(['ת"ז:', participant.id]).commit();

         // add headers
         worksheet.addRow(headers)

         // get first date
         let minimum = participant.startTime;
        
         if (participant.results != null){
         participant.results.forEach(result => {
            let category = result[0].value;

            for (let i = 1; i <= result.length; i++){
                let row = getRow(result[i], minimum, category);
                if (row != null){
                worksheet.addRow(row);
                }
            }
         });
        }
         
         worksheet.commit()
     });

    workbook.commit()
}

function getRow(result, minimum, categoty){
    if (result == null || minimum == null || result.dateOfChange == null){
        return null
    }

    let t1 = new Date(minimum);
    let t2 = new Date(result.dateOfChange);
    let timeFromStart = (t2-t1)/ 60000;

    return [result.value, timeFromStart, categoty];
}

function getMinimum(results){
    let dates = [];

    if (results == null){
        return 
    }

    results.forEach(result => {
        result.forEach(item => {
            if (item.dateOfChange != null){
                dates.push(item.dateOfChange)
            }
        })
    })

    minimum = _.min(dates);

    console.log(minimum)

    return minimum;
}