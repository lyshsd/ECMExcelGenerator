/**
 *	export-excel.js
 *	@author Patrick(patrick.holynova@gmail.com)
 */
;(function(doc, win) {
	// Check Compatibility
	var errormsg = 'Export Excel Tool: Please use IE 8+, FF 3.5+, Chrome, Saf or Opera';
	if(!doc.querySelectorAll) { console.info(errormsg); return; }
	
	var btntpl = '<span class="button-blue"><input type="button" id="xls_export_excel" value="Export Excel"></span>',
		formtpl = '<form style="display: none;" id="searchresulttoexcelform" method="post" enctype="multipart/form-data" action="http://9.115.144.33/ecm_extensions/exportexcel/generateexcel.php"><input name="contenttoexcel" id="contenttoexcel" /></form>';
		byId = function(id) {
			return doc.getElementById(id);
		},
		connect = function(node, /*String click*/event, callback) {
			if(doc.addEventListener) {
				return node.addEventListener(event, callback);
			} else if(doc.attachEvent) {
				return node.attachEvent('on' + event, callback);
			} else {
				console.info(errormsg);
			}
		},
		init = function() {
			var tbls = doc.querySelectorAll('table'), l = tbls.length;
				tbls[l-1].querySelectorAll('td')[0].innerHTML += btntpl,
				doc.querySelectorAll('#masthead')[0].innerHTML += formtpl;
			connect(byId('xls_export_excel'), 'click', function() {
				generate();
			});
		},
		generate = function() {
			var data = {};
			data.title = ['skey', 'uid', 'status', 'title', 'owner', 'version', 'modified'];
			data.items = [];
			var _getRowData = function(/*DOM Node*/trrow) {
				return [
					trrow.querySelectorAll('td[headers=c_select] input')[1].value,
					trrow.querySelectorAll('td[headers=c_select] input')[1].name.substr(4),
					trrow.querySelectorAll('td[headers=c_status] td')[1].innerHTML.replace(/^\s|\&nbsp;/g, ''),
					trrow.querySelectorAll('td[headers=c1] a')[1].innerHTML.replace(/'/g, "%%SQUOT%%").replace(/"/g, "%%DQUOT%%"),
					trrow.querySelectorAll('td[headers=c2]')[0].innerHTML,
					trrow.querySelectorAll('td[headers=c3]')[0].innerHTML,
					trrow.querySelectorAll('td[headers=c4]')[0].innerHTML
				];
			};
			var rows = doc.querySelectorAll('form[name=myForm] table tbody')[0].children,
				rowslen = rows.length;
			for(i = 0; i < rowslen; i++) {
				data.items.push(_getRowData(rows[i]));
			}
			byId('contenttoexcel').value = win.JSON.stringify(data);
			byId('searchresulttoexcelform').submit();
			byId('xls_export_excel').disabled = true;
			setTimeout(function() {
				byId('xls_export_excel').disabled = false;
			}, 8000);
	};
	init();
})(document, window);