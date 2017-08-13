// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window@Elements.js")

__H.register(__H.UI.Window.Elements,"Table","HTML Table",function Table(oTable){
	var d_tables = new ActiveXObject("Scripting.Dictionary")
	var table = null
	var thead = null
	var tbody = null
	var tfoot = null
	var trows = table
	
	this.isTable = function isTable(obj){
		return __HElem.isElements(obj) && obj.tagName == "TABLE"
	}
	
	this.isTBody = function isTBody(obj){
		return __HElem.isElements(obj) && obj.tagName == "TBODY"
	}
	
	this.setTable = function setTable(obj){
		if(this.isTBody(obj)) obj = obj.parentNode
		if(this.isTable(obj)){
			table = obj
			for(var i = table.childNodes.length-1; i >= 0; i--){
				if((table.childNodes(i).nodeName).toUpperCase() == "THEAD"){
					thead = table.childNodes(i)
				}
				else if(!tbody && (table.childNodes(i).nodeName).toUpperCase() == "TBODY"){
					tbody = table.childNodes(i)
				}
				else if((table.childNodes(i).nodeName).toUpperCase() == "TBODY"){
					tfoot = table.childNodes(i)
				}
			}
		}
		else {
			table = __H.byClone("table")
			tbody = __H.byClone("tbody")
			thead = __H.byClone("thead")
			tfoot = __H.byClone("tfoot")
			table.appendChild(thead)
			table.appendChild(tbody)
			table.appendChild(tfoot)
		}
		trows = tbody ? tbody : table
	}
	this.setTable(oTable)
	
	this.getTable = function getTable(){
		return table
	}
	
	this.createTable = function createTable(sOpt,sOpt2,sOpt3,sOpt4,sOpt5){
		sOpt4 = sOpt4 ? sOpt4 : ";"
		sOpt3 = sOpt3 ? sOpt3 : "\n"
		var r = sOpt2.split(sOpt3), s
		if(bSort){
			s = (r.slice(1,r.length)).sort()
			r = (new Array(r[0])).concat(s)
		}
		var t = new __H.Common.StringBuffer('<table width="98%" border="0" cellspacing="0" cellpadding="0" class="mba-table-2">')
		for(var i = 0, c, iLen = r.length; i < iLen; i++){
			c = r[i].split(sOpt4)
			t.append("\n<tr>")
			for(var j = 0, w, iLen2 = c.length; j < iLen2; j++){
				h = (c[j] == "" || c[j] == null) ? "&nbsp;" : c[j]
				w = j == 0 ? " width=\"1%\" nowrap" : ""
				t.append((i == 0 ? "<th" + w + " align=\"left\"> " + h + '</th>' : "<td" + w + "> " + h + '</td>'))
			}
			t.append("</tr>")
		}
		return t.append('</table>').toString()
	}
	/////////////////// ROW
	
	this.hasRows = function hasRows(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(isNaN(arguments[i]) || arguments[i] < 0  || trows.rows.length <= arguments[i]){
				return false
			}
		}
		return (i > 0)
	}
	
	this.getRow = function getRow(iRow){
		try{
			if(this.hasRows(iRow)){
				return trows.rows[iRow]
			}
		}
		catch(e){}
		return false
	}
	
	this.insertRow = function insertRow(iRow){
		iRow = !isNaN(iRow) && iRow >= 0 ? iRow : trows.rows.length
		var oRow = trows.insertRow(iRow)
		oRow.style.verticalAlign = "top";
		return oRow
	}
	
	this.deleteRow = function deleteRow(iRow){
		if(this.hasRows(iRow)){
			trows.deleteRow(iRow)
			table.refresh()
		}
	}
	
	this.deleteRows = function deleteRows(iRowStart,iRowEnd,bNotAll){
		if(typeof(iRowStart) != "number" || iRowStart < 0) iRowStart = 0
		if(typeof(iRowEnd) != "number" || iRowEnd < 0) iRowEnd = trows.rows.length-1
			
		iRowStart = iRowStart >= 0 && iRowStart > iRowEnd ? iRowEnd : iRowStart
		iRowEnd = iRowEnd > 0 && iRowEnd < iRowStart ? iRowStart : iRowEnd
		
		if(bNotAll && iRowStart == iRowEnd) return;
		for(var i = iRowEnd; i >= iRowStart; i--){
			trows.deleteRow(i)
		}
		table.refresh()
	}
	
	this.clearRows = function clearRows(){
		if(trows.rows.length > 0){
			for(var i = trows.rows.length-1; i >= 0; i--) trows.deleteRow(i);
		}
	}
	
	this.moveRow = function moveRow(iRowFrom,iRowTo){
		try{
			if(!this.hasRows(iRowFrom,iRowTo)) return false
			return trows.moveRow(iRowFrom,iRowTo);
		}
		catch(e){
			return false
		}
		finally{
			table.refresh();
		}
	}
	
	this.cloneRow = function cloneRow(iRow){
		if(trows.rows.length == 0) return;
		if(!this.hasRows(iRow)) iRow = trows.rows.length-1;
		//alert(trows.tagName)
		//alert("iRow="+iRow)
		var oClone = trows.rows(iRow).cloneNode(true);
		trows.rows(trows.rows.length-1).insertAdjacentElement("afterEnd",oClone);
	}
	
	this.getRowLength = function getRowLength(){
		return trows.rows.length
	}
	
	///////////////// CELL
	
	this.insertCell = function insertCell(oRow,sText){
		var oCell = oRow.insertCell()
		oCell.innerText = sText
		oCell.noWrap = true
	}
	
	this.setCell = function setCell(iRow,iCell,sText){
		if(!this.hasRows(iRow) || isNaN(iCell)) return;
		trows.rows[iRow].cells(iCell).innerText = sText
	}
	
	this.getCell = function getCell(iRow,iCell){
		if(!this.hasRows(iRow) || isNaN(iCell)) return;
		return trows.rows[iRow].cells(iCell)
	}
	
	this.addCell = function addCell(iRow,iCell){
		if(!this.hasRows(iRow) || isNaN(iCell)) return;
		trows.rows[iRow].insertCell(iCell)
		return true
	}
	
	this.delCell = function delCell(iRow,iCell){
		if(!this.hasRows(iRow) || isNaN(iCell)) return;
		trows.rows[iRow].deleteCell(iCell)
		table.refresh()
		return true
	}
	
	this.hasCell = function hasCell(oRow,iCell){
		try{
			return !isNaN(iCell) && iCell > 0 && oRow.cells(iCell)
		}
		catch(e){
			return false
		}
	}
	
	/////////////////// tbody
	
	this.getTbody = function getTbody(){
		return tbody
	}
	
	/////////////////// thead
	
	this.getThead = function getThead(){
		return thead
	}
	
	/////////////////// tbody
	
	this.getTfoot = function getTfoot(){
		return tfoot
	}
	
	this.refresh = function(){
		table.refresh()
	}
})


//// Global Define
var __HTable = new __H.UI.Window.Elements.Table()