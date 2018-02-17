var excel = require('node-xlsx').default;
var fs = require('fs');
var docxtemplater = require('docxtemplater')
var JSZip = require('jszip');

//display excel file
var excelData = excel.parse('Book1.xlsx');

var template = {};

excelData[0].data.forEach((el,i) => {
	i++;
	let x = 'ob'+i;
	let y = 're'+i;
	template[x] = el[0];
	template[y] = el[1];
});

//create docx template
var content = fs.readFileSync('templ.docx');
var zip = new JSZip(content);
var doc = new docxtemplater();
doc.loadZip(zip);

doc.setData(template);

try {
	doc.render();
}
catch (error) {
    var e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties,
    }
    console.log(JSON.stringify({error: e}));
    throw error;
}

var buf = doc.getZip().generate({type: 'nodebuffer'});
fs.writeFileSync('output2.docx', buf);