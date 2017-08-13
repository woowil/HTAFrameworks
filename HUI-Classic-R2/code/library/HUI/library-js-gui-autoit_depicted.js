// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-gui-autoit.js
Purpose:  Development script
Author:   Woody Wilson
Created:  08.07.2002
Version:  
Dependency:  library-js.js

Description:
JScript library scripting functions used for WSH files or HTA Applications.
 
Usage:

Description:
JScript library scripting functions used for WSH files or HTA Applications.

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

// Global Scripting Objects
try{
	if(!oFso || !oWsh || !oWno){
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oWsh = new ActiveXObject("WScript.Shell");
		var oWno = new ActiveXObject("WScript.Network");
	}
	
}
catch(ee){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
	var oShX = new ActiveXObject("Shell.Application");
}

var oAutX; // Must be loaded
var aut = new aut_object();

function aut_object(){
	this.t = 0	
}

function js_aut_setontop(sTitle,sText,iTop){
	try{
		if(typeof(oAutX) != "object") return;
		iTop = js_str_isnumber(iTop) ? iTop : 1 // 0=no, 1=yes
		sText = typeof(sText) == "string" ? sText : ""
		oAutX.WinSetOnTop(sTitle,sText,iTop)
		return true
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_aut_setstate(sOpt,sTitle,sText){
	try{
		if(typeof(oAutX) != "object") return;
		/*
		SW_HIDE = Hide window
		SW_SHOW = Shows a previously hidden window
		SW_MINIMIZE = Minimize window
		SW_MAXIMIZE = Maximize window
		SW_RESTORE = Undoes a window minimization or maximization
		*/
		try{
			sTitle = typeof(sTitle) == "string" ? sTitle : document.title
		}
		catch(ee){
			sTitle = "Untitled"
		}
		sText = typeof(sText) == "string" ? sText : ""
		oAutX.WinActivate(sTitle)
		if(sOpt == "minimize") oAutX.WinSetState(sTitle,sText,oAutX.SW_MINIMIZE)
		else if(sOpt == "maximize") oAutX.WinSetState(sTitle,sText,oAutX.SW_MAXIMIZE)
		else if(sOpt == "restore") oAutX.WinSetState(sTitle,sText,oAutX.SW_RESTORE)
		else if(sOpt == "show") oAutX.WinSetState(sTitle,sText,oAutX.SW_SHOW)
		else if(sOpt == "hide") oAutX.WinSetState(sTitle,sText,oAutX.SW_HIDE)
		else if(sOpt == "refresh"){ // free allocated memory
			//oAutX.WinSetState(sTitle,sText,oAutX.SW_HIDE)
			oAutX.WinSetState(sTitle,sText,oAutX.SW_MINIMIZE)
			oAutX.Sleep(1000)
			oAutX.WinSetState(sTitle,sText,oAutX.SW_RESTORE)			
			oAutX.WinSetState(sTitle,sText,oAutX.SW_SHOW)
			//oAutX.WinSetState(sTitle,sText,oAutX.SW_MAXIMIZE)
		}
		else return false
		return true
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}