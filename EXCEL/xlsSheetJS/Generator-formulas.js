var xlsx = require('xlsx');

//workBook class
function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

var exportBook = new Workbook();

var worksheet = {};

var cell = {f: 'A2+A3'};

var cellRef = xlsx.utils.encode_cell({r:0, c:0});

var range = {s:{r: 0, c: 0},
            e: {r: 10, c: 10}};



worksheet[cellRef] = cell;
worksheet['!ref'] = xlsx.utils.encode_range(range);

exportBook.SheetNames.push('test');
exportBook.Sheets.test = worksheet;


xlsx.writeFile(exportBook, 'formula sample.xlsx');