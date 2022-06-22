const excel = require('exceljs');
const csvToJson = require('csvtojson');
const fs = require('fs');
const excelToJson = require('convert-excel-to-json');

const workbook = new excel.Workbook();

//const filePath = 'optimizationStaticPivot';
const filePath = 'optimizationStaticPivot.xlsx';

const handleSheet = (sheet) => {
    
    console.log(sheet.name);

    for(let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(1);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };

    for(let rowIndex = 3; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(3);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
	};

    for(let rowIndex = 4; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(4);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };

    for(let rowIndex = 5; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(5);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 7; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(7);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 8; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(9);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 10; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(5);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 11; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(11);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 5; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(5);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 12; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(12);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
		for(let rowIndex = 13; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(13);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 14; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(14);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 15; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(15);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 16; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(16);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 17; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(17);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 18; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(18);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 19; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(19);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 20; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(20);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 5; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(5);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 21; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(21);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 5; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(5);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 22; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(22);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 23; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(23);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 24; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(24);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 25; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(25);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 26; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(26);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 27; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(27);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 28; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(28);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 29; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(29);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 30; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(30);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 31; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(31);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 32; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(32);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 33; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(33);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 34; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(34);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 35; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(35);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 36; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(36);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 37; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(37);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    };
	
	    for(let rowIndex = 38; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(38);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
	};

const onWorkbookOpenSuccess = () => {
    console.log('Открыт успешно.');

    workbook.worksheets.forEach(handleSheet);
	}
};


const onWorkbookOpenSuccess = () => {
    console.log('Открыт успешно.');

    workbook.worksheets.forEach(handleSheet);
};

const onWorkbookOpenError = reason => console.log(reason);

workbook.xlsx.readFile(filePath).then(onWorkbookOpenSuccess, onWorkbookOpenError);

var obj = new Object();
obj.name = "output-phase1";
var jsonAsString = JSON.stringify(obj);

//convert object to json string
var string = JSON.stringify(obj);

//convert string to Json Object
console.log(JSON.parse(string));

	console.log(JSON.parse(string));
const result = excelToJson({
    sourceFile: 'output-phase1.json',
    columnToKey: {
        A: 'Станция',
        C: 'Модель',
		D: 'Тип расчета',
		E: 'Пользователь',
		F: 'День',
    }

});
 