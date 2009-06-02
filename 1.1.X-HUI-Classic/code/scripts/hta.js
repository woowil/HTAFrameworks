// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     hta.js
Purpose:
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

	}
	catch(e){

	}
	finally{

	}
}


function hta_whoami(){
	try{
		var sCmd = "%comspec% /c echo " + (new Date()).formatDateTime() + "; " + oWno.UserDomain + "; " + oWno.UserName + "; " + pmt.ouser.FullName + "; " + oWno.ComputerName + "; " + oWsh.CurrentDirectory + "; " + oPMT.release + ">>" + pmt.fls.whoami
		oWsh.Run(sCmd,oReg.hide,false)
	}
	catch(e){
		alert("ERROR::hta_whoami(): " + e.description)
	}
}

function hta_onunload(){
	try{
		document.title = "..unloading"

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
	hta_onunload();
	window.close();
}

function hta_scroll(){
	try{

	}
	catch(e){}
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

function hta_test(sOpt){
	try{
		if(sOpt == "getfunctions"){
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
	}
	catch(e){}
}
