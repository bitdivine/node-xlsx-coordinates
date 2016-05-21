# node-xlsx-coordinates
Parse Excel coordinates:

	col2num('A') // -> 0
	col2num('Z') // -> 25
	col2num('AA') // -> 26

	parseCell('A4') // -> ['A4', 0, 3]
	parseCell('B5') // -> ['B5', 1, 4]

Also included is a function to convert an Excel sheet into a simple 2 dimensional array:

	tabulate( xlsx_sheet )


Example usage:

	var xlsx = require('xlsx');
	var workbook = xlsx.readFile(process.argv[2]);
	var sheetName = workbook.SheetNames[0];
	var sheet = workbook.Sheets[sheetName];
	tabulate(sheet).forEach(function(row){console.log(JSON.stringify(row));});

