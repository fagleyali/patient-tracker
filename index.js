const express =  require('express');
const Excel =  require ('exceljs');
const app = express();

app.use( (req,res) => {
    res.send('ok')
})

app.listen(3000, function(){
    console.log('Listening')
    
})

let workbook = new Excel.Workbook();
let fileName = 'Master Patient Tracker (2).xlsm (2).xlsm';

workbook.xlsx.readFile(fileName).then((result) => {
    let count = 0;
    let sheet1 = result.worksheets[1];
    let rowCount = sheet1.actualRowCount;
    let firstVisitColumn = sheet1.getColumn('O')
    let arr = []
    firstVisitColumn.eachCell( col => {
        if (new Date(col.value) > new Date('11/30/2019')  && col.value < new Date('01/01/2020')) {
            arr.push(col.row)
            
        }
        // console.log(new Date(col.value))
        
    })
    console.log(arr);
    console.log(new Date('11/30/2019'))

})
