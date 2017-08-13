// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-class-mgnt-storage.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-01
Version:
Dependency:  library-js.js, library-js__prototype.js

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
}


// Function _Class() must be loaded
class_mgnt_storage.prototype = new _Class("Storage","Class for managing disk and drives")

function class_mgnt_storage(oService,sComputer,sUser,sPass,bInit,bStatus,bDebug){
	try{
		this.wmi_cimv2 = oService
		this.regprov = null
		this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
		this.username = typeof(sUser) == "string" ? sUser : null
		this.password = typeof(sPass) == "string" ? sPass : null
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false
				
		
		if(bInit) this.init(oService)
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
