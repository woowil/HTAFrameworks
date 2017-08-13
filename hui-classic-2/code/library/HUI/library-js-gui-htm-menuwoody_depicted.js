// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

var oMyMenu = new htm_menu_object()

function htm_menu_object(){
	this.count = 0
	this.class_row_over = "cMenuItemOver"
	this.class_row_out = "cMenuItemOut"
	this.class_table = "cMenuTable"
}

function htm_menus(){
	try{
		var oMenu = htm_menu_create("test1",150)		
		htm_menu_elementrow("addrow",oMenu,"test1","Alt-1",null,false)
		/*
		htm_menu_elementrow("addrow",oMenu,"test2",null,null,true)
		htm_menu_elementrow("addrow",oMenu,"test3","Alt-3",null,true)
		htm_menu_elementrow("addrow",oMenu,"test4","Alt-2",null,false)
		htm_menu_elementrow("addrow",oMenu,"test5","Alt-5",null,false)
		*/
		htm_div_show(oMenu)
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_menu_create(sName,iWidth){
	try{
		 var oMenu = htm_menu_elementdiv("create",null,iWidth)		 
		 
		 document.body.appendChild(oMenu);
		 return oMenu
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_menu_elementdiv(sOpt,oDiv,iWidth){
	try{
		if(sOpt == "create"){
			iWidth = !isNaN(iWidth) ? iWidth : "auto"
			var oDiv = document.createElement("div")
			oDiv.id = "id" + ++oMyMenu.count;
			oDiv.className = "cMenuPop"
			oDiv.style.width = iWidth + "px";
			var oTable = htm_menu_elementtable("create")
			oDiv.appendChild(oTable)
			return oDiv
		}
		else if(sOpt == "over"){
			oDiv.setAttribute('className',oMyMenu.class_row_over)
		}
		else if(sOpt == "out"){
			oDiv.setAttribute('className',oMyMenu.class_row_out)
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_menu_elementtable(sOpt){
	try{
		 if(sOpt == "create"){
			var oTable = document.createElement("table")
			oTable.cellpadding = "0"
			oTable.cellspacing = "0"
			oTable.border = "0"
			oTable.width = "100%"
			oTable.align = "center"
			oRow = oTable.insertRow(0)
			oCell = oRow.insertCell()
			//oTable.className = oMyMenu.class_table 
			return oTable;
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

function htm_menu_elementrow(sOpt,oMenu,sText,sShortKey,sLink,bSubMenu,sIcon){
	try{
		 var oTable, oRow, oCell
		 if(sOpt == "addrow"){
			oTableOuter = oMenu.firstChild, len = oTableOuter.rows.length
			
			// Innner Table
			oTableInner = htm_menu_elementtable("create"), oTableInner.deleteRow(0)
			oRow = oTableInner.insertRow(0)			
			oCell = htm_menu_elementcell("create",oRow), oCell.width = "20px"
			oCell = htm_menu_elementcell("create",oRow,sText)
			oCell = htm_menu_elementcell("create",oRow), oCell.width = "35px", oCell.align = "right"
			if(sShortKey) oCell.innerText = sShortKey
			else if(bSubMenu) oCell.style.background = "#EEEEEE url(bin-pics/pmt_arrow2_right.gif) no-repeat fixed right"
			
			// Inner DIV
			oDivInner = document.createElement("div")
			oDivInner.onmouseover = new Function("htm_menu_elementdiv('over',this)")			
			oDivInner.onmouseout = new Function("htm_menu_elementdiv('out',this)")
			oDivInner.appendChild(oTableInner)
			
			oRows = oTableOuter.Rows(0), oCell = oRows.cell(0)
			oCell.appenChild(oDivInner)
			
			oTableOuter.refresh();
			return true;
		}
		else if(sOpt == "addseperator"){
			oCell = oRow.insertCell(), oCell.innerHTML = '<hr size="1" noshade width="100%">', oCell.colSpan = 3
		}
		else if(sOpt == "over"){
			oRow = oMenu
			oRow.setAttribute('className',oMyMenu.class_row_over)
		}
		else if(sOpt == "out"){
			oRow = oMenu
			oRow.setAttribute('className',oMyMenu.class_row_out)
		}
	}
	catch(e){
		htm_log_error(2,e);                    
		return false;
	}
}

function htm_menu_elementcell(sOpt,oRow,sText){
	try{
		if(sOpt == "create"){
			oCell = oRow.insertCell(), oCell.innerText = sText ? sText : " "
			//oCell.onmouseover = new Function("htm_menu_elementcell('over',this)")			
			//oCell.onmouseout = new Function("htm_menu_elementcell('out',this)")
			return oCell
		}
		else if(sOpt == "over"){
			oCell = oRow
			oCell.setAttribute('className',oMyMenu.class_row_over)
		}
		else if(sOpt == "out"){
			oCell = oRow
			oCell.setAttribute('className',oMyMenu.class_row_out)
		}
	}
	catch(e){
		htm_log_error(2,e);
		return false;
	}
}

