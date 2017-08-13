
// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-printer.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-01
Version:  
Dependency:  library-js.js, library-js-wmi.js

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

function js_printer_class(bStatus){
	try{
		this.cmd = "%comspec% /c rundll32 printui.dll,PrintUIEntry "
		
		this.Store = function(sPrinter){
			var sCmd = this.cmd + " /Ss /n \"" + sPrinter + "\" /a \"" + sPrinter + "-store-all.dat\""
		}
		
		this.StoreDevice = function(sPrinter){
			var sCmd = this.cmd + " /Ss /n \"" + sPrinter + "\" /a \"" + sPrinter + "-store-device.dat\" d"
		}
		
		this.StorePort = function(sPrinter){
			var sCmd = this.cmd + " /Ss /n \"" + sPrinter + "\" /a \"" + sPrinter + "-store-device.dat\" p"
		}
		
		this.ReStore = function(sPrinter){
			var sCmd = this.cmd + " /Sr /n \"" + sPrinter + "\" /a \"" + sPrinter + "-all.dat\" 2 7 c d g" // Without Security Descriptor
		}
		
		this.StoreAll = function(sPrinter){
			var sCmd = this.cmd + " /Sr /n \"" + sPrinter + "\" /a \"" + sPrinter + "-all.dat\""
		}
		
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
