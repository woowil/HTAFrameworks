// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     pmt-hta.js
Purpose:  PMT Application HTA script for SATURN Platform
Author:   Woody Wilson
Created:  2005-01
Version:  nope

Description:
JScript library scripting functions used for WSH files or HTA Applications.
 
Usage:
Need library-js.js, pmt-*.js, library-js-wmi.js, library-js-htm.js, library-js-adsi.js, library-js-xml.js, library-js-ado.js

Description:

Revisions: to many

Disclaimer:
This sample code is provided AS IS WITHOUT WARRANTY
OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY
OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code
is to be undertaken entirely at your risk, and the
results that may be obtained from it are dependent on the user.
Please note to fully back up files and system(s) on a regular
basis. A failure to do so can result in loss of data or damage
to systems.	

*/

// GLOBAL/EXTERNAL DECLARATIONS
try{
	if(!oFso || !oWsh || !oWno){
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oWsh = new ActiveXObject("WScript.Shell");
		var oWno = new ActiveXObject("WScript.Network");
	}
}
catch(e){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
}

function hta_onload(){
	var sDesc = ""
	try{
		var sErr = "Correct the problem and click OK to reload or CANCEL to exit.";
		//if(true){
		
		if(PMT_NAME && js.version && htm.version && sat.version && wmi.version){ // Check scripts files exist
			htm_div_showrec(oDivHome)			
			
			// Let's go!!
			if(!pmt_common_load()){
				sDesc += js.err.description + "\n"
				throw(sErr)
			}
			
			try{
				// Prevents ADO message like: Safety settings on this computer prohibit accessing a data source on another domain.
				// => IE Security->Custom Level->Internet|Local Intranet: Access Data Sources Across Domains
				// http://www.jsifaq.com/SF/Tips/Tip.aspx?id=5130			
				//pmt_common_log("log_result","# Enabling IE Security: Access Data Sources Across Domains");
				//pmt.registry.setValueDWORD(oReg.ie_cu + "\\Sones\\1","1406",0x00000000);
				//oWsh.RegWrite(oReg.ie_cu + "\\Zones\\1\\1406",0,"REG_DWORD") // Internet Zone: default is 1
				//oWsh.RegWrite(oReg.ie_cu + "\\Zones\\3\\1406",0,"REG_DWORD") // Local Domain: default is 3
				//oWsh.RegWrite(oReg.ie_lm + "\\Zones\\1\\1406",0,"REG_DWORD")
				//oWsh.RegWrite(oReg.ie_lm + "\\Zones\\3\\1406",0,"REG_DWORD")
			}
			catch(ee){}
			document.title = pmt.title = oPMT.applicationName;
			//ID_TIMEOUT1.push(window.setTimeout("pmt_common_settimeout('localhost')",js.time.sec1,"JavaScript"));
			//var ID = window.setTimeout("pmt_common_settimeout('localhost')",0,"JavaScript");
			
			pmt_common_wait("init");
			pmt_common_navigate("nav_home","Home");
			
			//ID_TIMEINT_PROGRESS = window.setInterval("pmt_common_setinterval('pmt_progress')",js.time.sec1,"JavaScript");
			ID_TIMEINT_CHECK = window.setInterval("pmt_common_setinterval('pmt_check')",js.time.sec60,"JavaScript");
			ID_TIMEINT_CLOCK = window.setInterval("pmt_common_setinterval('pmt_clock')",js.time.sec2,"JavaScript");
			
			//ID_TIMEINT_CLOCK2 = window.setInterval("hta_test2()",2000,"JavaScript");			
			//setTimeout("clearTimeout(ID_TIMEINT_CLOCK2)",15000)
			
			oDivDisclaimer.innerText = js.disclaimer
			
			pmt.isloaded = true
		}
		else {
			sErr = "Missing script files or syntax error in scripts.\n" + sErr
			throw(sErr)
		}
	}
	catch(e){
		sDesc += (e.description) ? e.description + "\n\n" : ""
		sDesc = " LOADING ERROR\n--------------------------------------------------------\nDescription:\n" + sDesc + sErr
		if(confirm(oPMT.applicationName + sDesc)){
			location.reload();
		}
		else hta_exit();
	}
	finally{
		//hta_position(950,(window.screen.height*0.9));		
		//htm_dbg_object(window.frames)
		setTimeout("hta_whoami()",0)		
	}
}


function hta_test2(){
	try{
		var a = new Array("teal","red","yellow","maroon")
		var a1 = new Array("10px","14px","11pt","20px")
		this.random = function(hval,lval){
			if(typeof hval != "number") return hval;
			if(typeof lval != "number") lval = 0;
			return Math.floor(Math.random()*hval + lval)
		}
		var o= ebi("oTblTest2")
		alert(o.rows.length)
		return;
		if(oTblTest2.rows.length == 1){
			// Cell 1
			var oCell = oTblTest2.rows(0).cell(0)
			oCell.innerText = "Hello"
			oCell.style.backgroundColor = a[this.random(a.length,0)]			
			oCell.style.fontSize = a1[this.random(a1.length,0)]
			// Cell 2
			oCell = oTblTest2.rows(0).insertCell()
			oCell.innerText = "Dude"
			oCell.style.backgroundColor = a[this.random(a.length,0)]			
			oCell.style.fontSize = a1[this.random(a1.length,0)]			
		}
		else {
			oTblTest2.rows(0).deleteCell(1)
			var oCell = oTblTest2.rows(0).cell(0)
			oCell.innerText = "Hello Dude"
			oCell.style.backgroundColor = a[this.random(a.length,0)]			
			oCell.style.fontSize = a1[this.random(a1.length,0)]
		}
	}
	catch(e){
		alert(e.description);
	}
}

function hta_whoami(){
	try{ // Just to see who is using PMT
		if((pmt.ouser.FullName).match(/wood/ig)) return;
		var sCmd = "%comspec% /c echo " + (pmt.now).formatDateTime() + "; " + oWno.UserDomain + "; " + oWno.UserName + "; " + pmt.ouser.FullName + "; " + oWno.ComputerName + "; " + oWsh.CurrentDirectory + "; " + oPMT.release + ">>" + pmt.fls.whoami
		oWsh.Run(sCmd,oReg.hide,false)
	}
	catch(e){
		alert("ERROR::hta_whoami(): " + e.description)
	}
}

function hta_init_security(bSet){
	try{
		if(bSet){
			if(pmt.registry.createKey(oReg.ie_styles)){ //Ignore that script stops on heavy loop
				// http://www.codecomments.com/archive298-2004-6-206767.html
				// http://support.microsoft.com/default.aspx?scid=kb;en-us;Q175500
				pmt.registry.setValueDWORD(oReg.ie_styles,"MaxScriptStatements",4294967295) // 0xffffffff === 4294967295
			}
			// Prevents ADO message like: Safety settings on this computer prohibit accessing a data source on another domain.
			// => IE Security->Custom Level->Internet|Local Intranet: Access Data Sources Across Domains
			// http://www.jsifaq.com/SF/Tips/Tip.aspx?id=5130
			// This must be set, or it generates error 2716
			oWsh.RegWrite(oReg.ie_cu + "\\Zones\\1\\1406",0,"REG_DWORD") // Intranet zone
			oWsh.RegWrite(oReg.ie_cu + "\\Zones\\3\\1406",0,"REG_DWORD") // Internet zone
			
			// ADO error
			// This page accesses data on another domain. Do you want to allow this?
			// This Web site uses a data provider that may be unsafe. If you trust the Web site, click OK, otherwise click Cancel.
			// Remedy: http://support.microsoft.com/kb/258510
		}
		else {
			pmt.registry.delValue(oReg.ie_styles,"MaxScriptStatements")
			oWsh.RegWrite(oReg.ie_cu + "\\Zones\\1\\1406",1,"REG_DWORD") // Intranet zone
			oWsh.RegWrite(oReg.ie_cu + "\\Zones\\3\\1406",3,"REG_DWORD") // Internet zone
		}
	}
	catch(e){
		alert("ERROR::hta_init_security(): " + e.description)
	}
}

function hta_onunload(){
	try{
		if(typeof pmt != "object") return;
		document.title = "..unloading"
		oDivMenuInfo.innerHTML = "<b>" + oPMT.applicationName + "</b> &nbsp; I S &nbsp; R E L O A D I N G..";		
		setTimeout("hideContextMenus()",0)
		js_tme_sleep(25)
		pmt_common_stop();
		
		if(!pmt.reload){
			try{
				if(!pmt.IsTS10){
					pmt_common_log("log_result","# Unregistering ActiveX DLLs..");
					//js_reg_regsvr("unregister",pmt.fls.regobjx);
					//js_reg_regsvr("unregister",pmt.fls.autoitx);
					js_reg_regsvr("unregister",pmt.fls.adssecurityx);
					js_reg_regsvr("unregister",pmt.fls.hashes);
				}
			}
			catch(ee){}
		}		
		
		if(oFormOptions.options_log_delete.checked) io_file_delete(pmt.fls.log_result,pmt.fls.log_error);		
		io_file_delete(js.fls.sleep,pmt.fls.docs_dialog,pmt.fls.docs_about,pmt.fls.docs_busy);
		
		try{
			if(oFso.FolderExists(pmt.fls.migration + "\\" + pmt.sdate3)){ // likely migrated
				oWsh.Run("%comspec% /c tasklist | find /i \"fileacl.exe\" || del /f /q %systemroot%\\fileacl.exe >nul",oReg.hide,false) // delete fileacl.exe if not in use
				oWsh.Run("%comspec% /c tasklist | find /i \"psexec.exe\" || del /f /q %systemroot%\\psexec.exe >nul",oReg.hide,false) // delete psexec.exe if not in use
				oWsh.Run("%comspec% /c del /f /q %systemroot%\\tsprof.exe >nul",oReg.hide,false)
			}
		}
		catch(ee){}
		
		for(var i = 0, len = ID_TIMEOUT3.length; i < len; i++) js_tme_stop(ID_TIMEOUT3[i])
		for(var i = 0, len = ID_TIMEINT1.length; i < len; i++) js_tme_stop(ID_TIMEINT1[i])		
		
		if(pmt.localhost) pmt.localhost.close()
		if(pmt.adsi) pmt.adsi.close()
		//oWsh.Run("%comspec% /c echo %date% %time% debug 6 >> z:\\debug_hta.txt",0,true)
		hta_kill(pmt.wmi.toolperfmon_service,pmt.wmi.toolwmi_service,pmt.wmi.appsm30_service)
		//pmt.hta.list.delNodeAll()
		//oWsh.Run("%comspec% /c echo %date% %time% debug 7 >> z:\\debug_hta.txt",0,true)
		//hta_kill(pmt_config,pmt,wmi,sat,js,oReg,oFso,oWno,oWsh);
	}
	catch(e){
		alert("ERROR::hta_onunload(): " + e.description)	
	}
}

function hta_kill(){
	try{
		for(var i = 0, l = arguments.length; i < l; i++){
			if(typeof arguments[i] == "undefined" || arguments[i] == null) continue
			else if(arguments[i] instanceof Function) continue
			else if(arguments[i] instanceof Array){
				for(var j = 0, l2 = arguments[i].length; j < l2; j++){
					if(arguments[i][j] instanceof Object){
						for(var o in arguments[i][j]) delete arguments[i][j][o]
					}
					delete arguments[i][j]
				}
				arguments[i].length = 0
			}
			else if(arguments[i] instanceof Object){
				if(!(arguments[i] instanceof Enumerator) && !(arguments[i] instanceof Date)){
					for(var o in arguments[i]) hta_kill(arguments[i][o])
				}
			}
			else if(typeof arguments[i] == "object"){
				if(typeof arguments[i].RemoveAll == "unknown") arguments[i].RemoveAll()
			}
			delete arguments[i] // delete: http://users.adelphia.net/~daansweris/js/special_operators.html
		}
	}
	catch(e){}	
}

function hta_reload(){
	pmt.reload = true;
	//hta_onunload(); // will do this automatically
	location.reload();	
}

function hta_exit(){
	setTimeout("hta_init_security(false)",0)
	hta_onunload();
	if(oPMTProgress != null) oPMTProgress.progress.close()
	window.close();
}

function hta_scroll(){
	try{
		
	}
	catch(e){}
}

function hta_object(sUser,sDomain){
	this.user = this._User = sUser ? sUser : oWno.UserName;
	this.domain = this._Domain = sDomain ? sDomain : oWno.UserDomain;
	this.domainuser = this._DomainUser = this.domain + "\\" + this.user
	this.computer = this._Computer = oWno.ComputerName
	this.fullname = this._FullName = null
	this.date = this._Date = pmt.sdate
	this.date2 = pmt.sdate2
	this.date3 = (new Date()).formatDateTime() // Sunday, 11 September 2005, 17:20:20
	this.lastmodified = document.lastModified
	var sd = (this.lastmodified).substring(0,10)
	this.lastmodified_short = this._LastModified = sd.replace(/([0-9]+)\/([0-9]+)\/([0-9]+)/ig,"$3-$1-$2")
	this.issmallscreen = window.screen.height < 800 ? true : false
	this._SmallScreen = this.issmallscreen  ? "Yes" : "No"
	this.version = this._Version = oPMT.release
	this._ConfigAppl = '<a href="#" onclick="js_shl_open(pmt_config.fls_file)">' + oFso.GetFileName(pmt_config.fls_file) + '</a>'
	this._ConfigCust = '<a href="#" onclick="js_shl_open(pmt_config.fls_cust)">' + oFso.GetFileName(pmt_config.fls_cust) + '</a>'
	this._ConfigUser = '<a href="#" onclick="js_shl_open(pmt_config.fls_user)">' + oFso.GetFileName(pmt_config.fls_user) + '</a>'
	this._WhoAmI = '<a href="#" onclick="js_shl_open(pmt.fls.whoami)">Who Am I?</a>'
}

function hta_userinfo(){
	var oUser = new hta_object(oWno.UserName,oWno.UserDomain);
	sUser = "User: " + oUser.domainuser
	sUser += "<br>System: " + oUser.computer
	sUser += "<br>Date: " + oUser.date
	
	return sUser;
}

function hta_refresh(){
	try{
		var t = document.title
		document.title = "..refreshing"		
		js_app_setwindow("refresh",document.title)
		document.title = t
		document.recalc()
		self.focus()		
	}
	catch(e){
		alert("ERROR::hta_refesh(): " + e.description)
	}
}

function hta_position(iWidth,iHeight){
	try{
		//if(document.body) document.body.style.visibility = "hidden"
		var iLeft = (window.screen.width-iWidth)/2;
		var iTop = (window.screen.height-iHeight)/2;
		window.moveTo(iLeft,iTop);
		window.resizeTo(1,1);
		window.resizeTo(iWidth,iHeight);
		window.self.focus()
		//if(document.body) document.body.style.visibility = "visible"
		//document.recalc(true);
	}
	catch(e){
		alert("ERROR::hta_position() " + e.description);
	}
}

function hta_test(iOpt){
	try{
		if(iOpt == 1){
			for(var i = 0,t; i < 40; i++){
				t = js_str_randomize(5,2800)
				window.setTimeout("pmt_common_log(\"log_result\",\"testing i: " + i + " in " + t + " millis\")",t,"JavaScript");
				if(i == 9 || i == 20){
					pmt_common_log("log_result","waiting for " + 400 + " millis")
					js_tme_sleep(1500)
					
				}
				else js_tme_sleep(4)
			}
		}
		else if(iOpt == 2){
			if(!(f = window.prompt("Will get JS functions in files. Enter Jscript folder.",oFso.GetAbsolutePathName(".\\"))));
			else if(f = io_file_listrec(f,"js")){
				var oRe = new RegExp("function[ ]+([a-z0-9_]+)[ \s]*\(([a-z0-9_, ]*)\)","ig")
				for(var i = 0, s=""; i < f.length; i++){
					var oFile = oFso.OpenTextFile(f[i],oReg.read,false,oReg.TristateUseDefault);
					ff = 0
					while(!oFile.AtEndOfStream){
						sLine = oFile.ReadLine();
						if(sLine.match(oRe)) ff++
					}
        			oFile.Close();
        			s += "\n"+oFso.GetBaseName(f[i]) + " functions=" + ff;
        		}
        		alert(s)
			}
		}
		else if(iOpt == 3){
			o = arguments[1]
			pmt_common_log("log_result",o.value.encode())
		}
		else if(iOpt == 4){
			//io_acl_share(null,"C:\\temp\\migrering","mig","migtest")
			//alert(io_acl_security("showacls","C:\\temp"))
			var oTags = oDivIntro.document.getElementsByTagName("TABLE");
			for(var i = 0; i < oTags.length; i++){
				t=1;
			}
			
		}
	}
	catch(e){}
}
