(async function() {
    const excel = require('exceljs');
    const workbook = new excel.Workbook();
    // use readFile for testing purpose
    // await workbook.xlsx.load(objDescExcel.buffer);
    await workbook.xlsx.readFile(process.argv[2]);
    let jsonData = [];
    workbook.worksheets.forEach(function(sheet) {
        // read first row as data keys
        let firstRow = sheet.getRow(1);
        if (!firstRow.cellCount) return;
        let keys = firstRow.values;
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber == 1) return;
            let values = row.values
            let obj = {};
            for (let i = 1; i < keys.length; i ++) {
                obj[keys[i]] = values[i];
            }
            jsonData.push(obj);
        })

    });
    console.log(jsonData);
})();