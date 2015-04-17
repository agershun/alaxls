// main function {
//     var grid;    
//     arrtoob(arraytest);   //arraytest is an multi-dimensional array which I created earlier. It's about 25x45 (25 wide, 45 tall) with some empty cells
//     exportxlsx();
// }

var branch = "neverland";
// Create demo column
var columns = [];
for(var j=0;j<40;j++){
	columns.push({columnid:j, name: 'Column #'+j, type:'numeric', width:120});
};

var data = [];
for(var i=0;i<50;i++) {
	var r = {};
	for(var j=0;j<columns.length;j++) {
		// Fill array with some data
		r[columns[j].columnid] = i*100+j;
	}
	data.push(r);
}


// function arrtoob(arr){
// 	var data = []; 
// 	arr.forEach(function(row){
// 		var r = {}; 
// 		row.forEach(function(v,idx){ 
// 			r[idx]=v
// 		}); 
// 		data.push(r);
// 	});
// 	var columns = []; 
// 	for(var i=0;i<arr[0].length;i++){
// 		columns.push({columnid:i,type:'numeric'})
// 	};
// 	grid={data:data,columns:columns};
// 	}
    
window.exportxlsx = function exportxlsx() {
	window.gridExportToExcel = (function () {
		var a = document.createElement("a");
		document.body.appendChild(a);
		a.style = "display: none";
			var s = generateExcel();
			var blob = new Blob([s], { type: 'application/vnd.ms-excel' });
			url = window.URL.createObjectURL(blob);
			a.href = url;
			a.download = branch+".xls";
			a.click();
			window.URL.revokeObjectURL(url);
	}());
}

function generateExcel() { 
	if (typeof title == "undefined") title = "Sheet1";

	var s = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" \
	xmlns="http://www.w3.org/TR/REC-html40"><head> \
	<meta charset="utf-8" /> \
	<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets> ';

	s+=' <x:ExcelWorksheet><x:Name>' + title + '</x:Name><x:WorksheetOptions><x:DisplayGridlines/>     </x:WorksheetOptions> \
	</x:ExcelWorksheet>';

	s += '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>';

	s += '<colgroups>';
	columns.forEach(function (col) {
		s += '<col style="width: '+col.width+'px"></col>';
	});

	s += '<thead><tr>';
	columns.forEach(function (col) {
		s += '<th style="background-color: #E5E5E5; border: 1px solid black;">' + col.name + '</th>';
	});

	s += '<tbody>';
	data.forEach(function(d){
		s += '<tr>';

		columns.forEach(function (col) {
			var value = d[col.columnid];
console.log(value);
			s += '<td ';
			if (col.kindid == "money") {
			   s += "style = 'mso-number-format:\"\\#\\,\\#\\#0\\\\ _р_\\.\";white-space:normal;'"; 
			} else if (col.type == "numeric") s += "";//" style = 'mso-number-format:\"\\@\";'";
			else if (col.kindid == "date") s += " style='mso-number-format:\"Short Date\";'";
			else s += " style='mso-number-format:\"\\@\";'";
			s += '>';
			if(typeof value == "undefined") { 
				s += ''; 
			} else if (col.type == "numeric") {
				s += value.toString(); 
			} else if (col.type == "date") {
//				s += moment(value).format('DD.MM.YY');
                s += value.toString();
			} else if (col.type == "money") {
				s += value.toFixed(2);
			} else s += d[col.id];
            s += '</td>';
		});
	});
	s += '</table></body></html>';

	return s;
}
