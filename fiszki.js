var excel = require('node-xlsx').default;
var fs = require('fs');

var data = fs.readFile('Book1.xlsx');

var excelData = excel.parse('Book1.xlsx');
console.log(excelData[0].data);
