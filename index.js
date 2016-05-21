/* A->0, B->1, Z->25, AA->26 etc. */
function col2num(s){ var ans = 0; for (var i=0; i<s.length; i++) ans = ans*26+s.charCodeAt(i)-64; return ans - 1;}

/* colum index -> col label strlen e.g. AA == 26 -> 2 */
function digits(x){return Math.floor(Math.log((x+1)*25+1)/Math.log(26));}

/* 0 -> A, 25 -> Z, 26 -> AA ... */
function num2col(n){ var ans = new Array(digits(n)).fill(0); n -= ((Math.pow(26,ans.length)-1)/25) - 1; for(var i=ans.length; i--;){ ans[i] = n%26; n = (n - (n%26))/26; } return ans.map(function(n){ return String.fromCharCode(n+65); }).join(''); }

/* A4 -> [0, 3, 'A4'] -- Note that the 1-up row count in the label is changed to a 0-up index. */
function parseCell(s){ var m = s.match(/([A-Z]+)([0-9]+)/); if (m !== null) return [col2num(m[1]),Number(m[2])-1,m[0]]; }

/* [0,0] -> A1  ... [x,y] -> cell_key */
function stringifyCell(cell) { return num2col(cell[0])+(1+cell[1]); }

/* sheet -> [row1, row2, ...] where row = [cellvalue, cellvalue, ...] */
function tabulate(sheet){
	var rows = [];
	Object.keys(sheet).map(parseCell).filter(function(x){return x!==undefined;}).forEach(function(parsed_key){
		var key = parsed_key[2]
		  , col = parsed_key[0]
		  , row = parsed_key[1]
		  , val = sheet[key].v;
		if (undefined === rows[row]) rows[row] = [];
		rows[row][col] = val;
	});     
	return rows;
}       

module.exports =
{ col2num:	col2num
, num2col:	num2col
, parseCell:	parseCell
, stringifyCell:stringifyCell
, tabulate:	tabulate
};

if (!module.parent) {
	var xlsx = require('xlsx');
	var workbook = xlsx.readFile(process.argv[2]);
	var sheetName = workbook.SheetNames[0];
	var sheet = workbook.Sheets[sheetName];
	tabulate(sheet).forEach(function(row){console.log(JSON.stringify(row));});
}
