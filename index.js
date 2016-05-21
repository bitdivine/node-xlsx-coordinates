/* A->0, B->1, Z->25, AA->26 etc. */
function col2num(s){ var ans = 0; for (var i=0; i<s.length; i++) ans = ans*26+s.charCodeAt(i)-65; return ans + ((Math.pow(26,s.length)-1)/25) - 1;}

/* A4 -> ['A4', 0, 3] -- Note that the 1-up row count in the label is changed to a 0-up index. */
function parseCell(s){ var m = s.match(/([A-Z]+)([0-9]+)/); if (m !== null) return [m[0],col2num(m[1]),Number(m[2])-1]; }

/* sheet -> [row1, row2, ...] where row = [cellvalue, cellvalue, ...] */
function tabulate(sheet){
	var rows = [];
	Object.keys(sheet).map(parseCell).filter(function(x){return x!==undefined;}).forEach(function(parsed_key){
		var key = parsed_key[0]
		  , col = parsed_key[1]
		  , row = parsed_key[2]
		  , val = sheet[key].v;
		if (undefined === rows[row]) rows[row] = [];
		rows[row][col] = val;
	});     
	return rows;
}       

module.exports =
{ col2num:	col2num
, parseCell:	parseCell
, tabulate:	tabulate
};

if (!module.parent) {
	var xlsx = require('xlsx');
	var workbook = xlsx.readFile(process.argv[2]);
	var sheetName = workbook.SheetNames[0];
	var sheet = workbook.Sheets[sheetName];
	tabulate(sheet).forEach(function(row){console.log(JSON.stringify(row));});
}
