/**
 *	export-excel.js
 *	@author patrick.holynova@gmail.com
 */
;(function(doc, win) {
	// Check Compatibility
	var errormsg = 'Export Excel Tool: Please use IE 8+, FF 3.5+, Chrome, Saf or Opera';
	if(!doc.querySelectorAll) { console.info(errormsg); return; }
	
	var overlaytpl = '<div id="searchresulttoexcel" style="display:none;position:fixed;top:0px;left:0px;z-index:949;width:100%;height:100%;background-color:rgba(0, 0, 0, 0.5);"><div style="position:absolute;background-color:#fff;width:390px;height:300px;top:50%;left: 50%;margin-top:-200px;margin-left:-180px;border-radius:3px;box-shadow:0 0 10px 0 rgba(255,255,255,0.7)"><h2>Export Excel</h2><p id="xsl_category"><input type="checkbox" id="xls_skey" name="xls_skey" value="skey" /><label for="xls_skey">Synkey</label><input type="checkbox" id="xls_uid" name="xls_uid" value="uid" /><label for="xls_uid">Unique id</label><input type="checkbox" id="xls_status" name="xls_status" value="status" /><label for="xls_status">Status</label><input type="checkbox" id="xls_title" name="xls_title" value="title" /><label for="xls_title">Title</label><input type="checkbox" id="xls_owner" name="xls_owner" value="owner" /><label for="xls_owner">Owner</label><input type="checkbox" id="xls_version" name="xls_version" value="version" /><label for="xls_version">Version Created</label><input type="checkbox" id="xls_modified" name="xls_modified" value="modified" /><label for="xls_modified">Last Modified Date</label></p><p><input type="button" value="Export" id="xls_export"><input type="button" value="Cancel" id="xls_cancel"></p><form style="display: none;" id="searchresulttoexcelform" method="post" enctype="multipart/form-data" action="http://9.115.144.33/ecm_extensions/exportexcel/generateexcel.php"><textarea name="contenttoexcel" id="contenttoexcel"></textarea></form></div></div>',
		btntpl = '<span class="button-blue"><input type="button" id="xls_export_excel" value="Export Excel"></span>',
		byId = function(id) {
			return doc.getElementById(id);
		},
		addClass = function(node, newClassName) {
			if(node.className.indexOf(newClassName) > -1) { return; }
			else {
				node.className += ' ' + newClassName;
			}
		},
		removeClass = function(node, oldClassName) {
			var reg = new RegExp(oldClassName, 'g');
			return node.replace(reg, '');
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
		addScript = function(url) {
			var ele = doc.createElement('script');
			ele.src = url;
			ele.setAttribute("type","text/javascript");
			doc.getElementsByTagName('head')[0].appendChild(ele);
		},
		initBtn = function(btntpl) {
			var tbls = doc.querySelectorAll('table'), l = tbls.length;
			tbls[l-1].querySelectorAll('td')[0].innerHTML += btntpl;
			connect(byId('xls_export_excel'), 'click', function() {
				show('searchresulttoexcel');
			});
		},
		collectData = function() {
			console.log("collect data begin...");
			var data = {};
			data.title = ['skey', 'uid', 'status', 'title', 'owner', 'version', 'modified'];
			data.items = [];
			//pre = window.location.protocol + '//' + window.location.host;
			var _getRowData = function(/*DOM Node*/trrow) {
				// skey, uid, status, title, owner, version, modified
				// Speed issue
				return [
					trrow.querySelectorAll('td[headers=c_select] input')[1].value,
					trrow.querySelectorAll('td[headers=c_select] input')[1].name.substr(4),
					trrow.querySelectorAll('td[headers=c_status] td')[1].innerHTML.replace(/^\s|\&nbsp;/g, ''),
					trrow.querySelectorAll('td[headers=c1] a')[1].innerHTML,
					/*pre + trrow.querySelectorAll('td[headers=c1] a')[1].href,*/
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
			console.log(data)
			byId('contenttoexcel').innerHTML = win.JSON.stringify(data);
			byId('searchresulttoexcelform').submit();
		},
		initOverlay = function(overlaytpl) {
			var wrapper = doc.createElement('div');
			wrapper.innerHTML = overlaytpl;
			doc.querySelectorAll('body')[0].appendChild(wrapper.firstChild);
			connect(byId('xls_cancel'), 'click', function() {
				hide('searchresulttoexcel');
			});
			connect(byId('xls_export'), 'click', collectData);
		},
		show = function(id) {
			var overlay = byId(id);
			if(!overlay) { return; }
			else {
				doc.getElementById(id).style.display = 'inline';
			}
		},
		hide = function(id) {
			var overlay = byId(id);
			if(!overlay) { return; }
			else {
				doc.getElementById(id).style.display = 'none';
			}
		},
		init = function() {
			initBtn(btntpl);
			initOverlay(overlaytpl);

		};
	// Init
	init();
})(document, window);