const Excel = require("exceljs");
const workbook = new Excel.Workbook();
const workbook_write = new Excel.Workbook();
const filePathreading = `so_diem_tong_ket_khoi_khoi_12.xlsx`
const filepathwriting = `bangdiemfilter.xlsx`

var CONFIG = require('./config.json');
// console.log(CONFIG);
// return;
workbook.xlsx.readFile(filePathreading).then(() => {
  // console.log("Sheet name---", workbook.getWorksheet(1).name)
  // console.log("Sheet name---", workbook.getWorksheet(2).name)

  // workbook.worksheets.forEach((sheet, i) => {
  //   // console.log("Array index:", i)
  //   // console.log("exceljs id:", sheet.id)
  //   // console.log("Sheet name---", sheet.name)
  //   list_worksheets.push(sheet.name)
  // })

  let excel_data = [];
  let list_worksheets = getListWorkSheet(workbook)
  let columns = getColumnName( workbook.getWorksheet(list_worksheets[0]))

  list_worksheets.forEach( sheet =>{
    let worksheet = workbook.getWorksheet(sheet);
    let data = getData(worksheet,columns);

    let column_condition = CONFIG.sheets.find( sheet_item => sheet_item.sheet_name == sheet ).column_condition;
    
 
    data = data.filter(data => data[column_condition] > 8.5)
    data = data.filter(data => data["Học lực"] == 'Giỏi' ||  data["Học lực"] == 'Tốt' )
    data = data.filter(data =>  data["Hạnh kiểm"] == 'Tốt' )

    data = data.sort(function(a, b)  {
      return - a[column_condition] + b[column_condition]
    })


    excel_data.push({
      sheet_name:sheet,
      data:data,
    })

  })

    excel_data.forEach(sheet =>{
      const worksheet = workbook_write.addWorksheet(sheet.sheet_name, {
      });
      worksheet.addRow(columns) 
      sheet.data.forEach(data=>{
        let row = [];

        columns.forEach((value) => {

          row.push(data[value])
        })
        worksheet.addRow(row) 
        
      })


      // worksheet.addRow(columns);
    })

    workbook_write.xlsx.writeFile('export.xlsx');
 
  

});




function getListWorkSheet(workbook){
  let list_worksheets = []
  workbook.worksheets.forEach((sheet, i) => {
    // console.log("Array index:", i)
    // console.log("exceljs id:", sheet.id)
    // console.log("Sheet name---", sheet.name)
    list_worksheets.push(sheet.name)
  })
  return list_worksheets;
}

function getColumnName(worksheet){
  let columns = []
  worksheet.eachRow(function(row, rowNumber) {
    if(rowNumber == 6){
      row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
        if(cell.value !== null){
          columns.push(cell.value)
        }
      });
    }
  });
  return columns
}

function getData(worksheet, columns){
  let data = [];
  worksheet.eachRow(function(row, rowNumber) {

    if(rowNumber >= 8 && row.getCell(1).value !== null ){
      let row_data = [];
      row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
        row_data[`${columns[colNumber-1]}`] = cell.value
      });
      data.push(row_data);
    }
    
  });
  return data;
}
