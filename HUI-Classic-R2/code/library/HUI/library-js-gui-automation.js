// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-gui-automation.js
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
	if(!oFso || !oWsh || !oWno || !oXml){
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

////////////////////////////////////////////////////////////////////////////////////////////////
/////// AUTOMATION FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////


function js_app_delaykey(sKey,iDelay){
	try{
		iDelay = js_str_isnumber(iDelay) ? iDelay : 100 
		js_tme_sleep(iDelay);
		oWsh.SendKeys(sKey);
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_app_activate(sTitle,iDelay,iTime){
	try{
		if(typeof(sTitle) != "string") return false;
		iTime = iTime ? iTime : 60000;
		iDelay = iDelay ? iDelay : 400;
		js_tme_sleep(iDelay);
		for(var idle = 0; !oWsh.AppActivate(sTitle); idle += 200){
			if(idle > iTime){
				var msg = "Error (Time out):\n\nThe window '" + sTitle + "' couldn't be found. Check the function caller for debugging.\nFunction Caller:\n\n" + js_app_activate.caller;
				throw(msg);
				break;
			}
			js_tme_sleep(200);
		}
		oWsh.AppActivate(sTitle);
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_app_close(sTitle,iDelay,iTime){
	try{
		if(typeof(sTitle) != "string") return false;
		iTime = js_str_isnumber(iTime) ? iTime : 10000;
		iDelay = js_str_isnumber(iDelay) ? iDelay : 400;
		js_tme_sleep(iDelay);
		for(var idle = 0; !oWsh.AppActivate(sTitle); idle += 200){
			if(idle > iTime){
				break;
			}
			js_tme_sleep(200)
		}
		if(oWsh.AppActivate(sTitle)){
			js_app_delaykey("%{F4}");
			if(oWsh.AppActivate(sTitle)){ // DOS windows
				oWsh.SendKeys("% "); // ALT SPACE
				if(oWsh.AppActivate(sTitle)){ // DOS windows running tail
					oWsh.SendKeys("^C"); // CTRL-C
				}
			}
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_app_setwindow(sOpt,sTitle){
	try{
		if(typeof(sTitle) != "string") return;
		if(sOpt == "minimize"){
			js_app_activate(sTitle)
			oWsh.SendKeys("% ")
			oWsh.SendKeys("n")
		}
		else if(sOpt == "refresh"){
			js_app_activate(sTitle,100,1000)
			oWsh.SendKeys("% ")
			oWsh.SendKeys("n")
			js_tme_sleep(50)
			js_app_activate(sTitle,100,1000)
			oWsh.SendKeys("% ")
			oWsh.SendKeys("r")
		}
		/*
		switch(iOpt){
			case 1 : { //  window
				js_app_activate(sTitle);
				for(var i = 2; i < arguments.length; i++){
					js_app_delaykey(arguments[i]);
				}
				break;
			}
			case 2 : case 3 : case 4 : {
				if(iOpt == 2) sKey = (sLang == "en") ? "% n" : "% i"; // Minimize
				else if(iOpt == 3) sKey = (sLang == "en") ? "% x" : "% m"; // Maximixe
				else sKey = (sLang == "en") ? "% r" : "% s"; // Restore				
				js_app_activate(sTitle,100);
				js_app_delaykey(sKey);
				break;
			}                                  
			case 5 : {// Minimize/Maximize All
			
				break;
			}                  
			case 6 : {// Get window title (Requires 'JSSys3.Ops')
				if(js.jssysx){
					var sTitle = "", aWindow = new Array(), aResult = new Array();
					while(true){
						sTitle = js.jssysx.GetActiveWindowTitles(sTitle);
						if(aWindow[sTitle]) break;
						aWindow[sTitle] = true;
						aResult.push(sTitle);
						oWsh.SendKeys("%{TAB}");
					}
					bResult = aResult;
				}
				else bResult = false;
				break;
			}
			default: {
				bResult = false;
				break;
			}
		}
		*/
		return bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_registry(sComputer){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName
		oWsh.Run("%comspec% /c regedit.exe",oReg.hide,false);
		js_app_activate("Registry Editor",1200);
		js_app_delaykey("%F");
		js_app_delaykey("C");
		js_app_delaykey(sComputer);
		js_app_delaykey("{ENTER}")
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_rebootNO(){
	try{
		js_app_sendKey("^{ESC}");
		js_app_sendKey("{UP}");
		js_app_sendKey("{ENTER}");
		js_app_activate("Avslutt Windows");
		js_app_sendKey("%t");
		js_app_sendKey("%J");
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_app_winzip81(sFile){
	try{
		if(oFso.FileExists(sFile)){
			oWsh.Run("%comspec% /c " + sFile + " /noqp /autoinstall",0,false);
			js_app_activate("WinZip 8.1 SR-1 Setup",5000);
			js_app_delaykey("%S");
			js_app_activate("WinZip Setup",5000);
			js_app_delaykey("{ENTER}");
			js_app_activate("WinZip Setup",1500);
			js_app_delaykey("%N");
			js_app_activate("License Agreement and Warranty Disclaimer",1500);
			js_app_delaykey("%Y");
			js_app_activate("WinZip Setup",500);
			js_app_delaykey("%N");
			js_app_activate("WinZip Setup",500);
			js_app_delaykey("%C%N");
			js_app_activate("WinZip Setup",500);
			js_app_delaykey("%C%N");
			js_app_activate("WinZip Setup",500);
			js_app_delaykey("%S%d%N");
			js_app_activate("WinZip Setup",500);
			js_app_delaykey("%i%N");
			js_app_delaykey("{ENTER}");
			js_app_activate("WinZip Tip of the Day",2500);
			js_app_delaykey("%C");
			js_app_delaykey("%{F4}");			
			
			return true;
		}
		else return false;
	}
	catch(e){
		js_log_error(2,e);
		return false
	}
}

function js_app_ar505enu(sFile){
	try{
		if(oFso.FileExists(sFile)){
			oWsh.Run("%comspec% /c " + sFile,0,false);
			js_app_activate("Acrobat Reader 5.0.5 Setup",20000);
			js_app_delaykey("%N");
			js_app_activate("Choose Destination Location",1500);
			js_app_delaykey("%N");
			js_app_activate("Information",10000);
			js_app_delaykey("{ENTER}");	
			return true;
		}
		else return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_app_officewinword(iOpt,oActiveX,sFilter,sDir,oTextErr){
	try{
		if(iOpt == 1 || iOpt == 2){
			var sFile;
			var sWd = oActiveX ? oActiveX : "Word.Application.8"; // MS Office 97
			//var sWd = oActiveX ? oActiveX : "Word.Document"; // MS Office 2002
			var oWd = new ActiveXObject(sWd);
			var sWdStateMinimize = 2;
			var sWdDialogFileOpen = 0x50;
			var sWdDialogFileSaveAs = 0x54;
			var sWdDialogFileOpenMultiple = 0x100;
			var sWdInitialDir = sDir ? sDir : "c:\\";
			var sWdFileFilter = sFilter ? sFilter : "*.*";
			oWd.Visible = false; // Hides Word
			oWd.WindowState = sWdStateMinimize; // Keeps Word from popping up, if/when you open a file
			oWd.ChangeFileOpenDirectory(sWdInitialDir);	
		}
		if(iOpt == 1){ // Open Dialog
			var oWdOpen = oWd.Dialogs(sWdDialogFileOpenMultiple);
			oWdOpen.Name = sWdFileFilter;
			sWdDlgBtnClicked = oWdOpen.Show(); // Show dialog open file, return button clicked
			//sWdDlgBtnClicked = oWdOpen.Display() // show dialog, Without opening the file
			switch(sWdDlgBtnClicked){
				case '-2' : sResult = "Close"; break;
				case '-1' : sResult = "OK"; break;
				case '0' : sResult = "Cancel"; break;
				default: sResult = "Button number: " + sWdDlgBtnClicked; break;
			}
			sFile = oWdOpen.Name;
		}
		else if(iOpt == 2){ // SaveAs dialog
			var oWdOpen = oWd.Dialogs(sWdDialogFileSaveAs); // Note: savesa won't work unless you have a file open 
			oWdOpen.Name = "foobar";
			sWdDlgBtnClicked = oWdOpen.Display(); // show dialog, Without opening the file
			sFile = oWdOpen.Name; // the filename selected
		}
		else return false;
	}
	catch(e){
		js_log_error(2,e,oTextErr);
		return false;
	}
	finally{
		if(oWd){
			oWd.Application.Quit(false); // Quits and save changes
			oWd = null;
		}
		return sFile;
	}
}