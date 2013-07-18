// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     library-js-io-drive.js
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

function js_io_driveinfo(sDrive){
	try{
		var Obj = new Object();
		var oDrive = oFso.GetDrive(oFso.GetDriveName(sDrive))
		switch(oDrive.DriveType){
			case 0: Obj.type = "unknown"; break;
			case 1: Obj.type = "removable"; break;
			case 2: Obj.type = "fixed"; break; // Both regular drive and substitute drives
			case 3: Obj.type = "network"; break;
			case 4: Obj.type = "cdrom"; break;
			case 5: Obj.type = "ramdisk"; break;
			default : Obj.type = null; break;
		}
		if(Obj.type != null){
			Obj.drvLetter = oDrive.DriveLetter;
			Obj.isready = oDrive.IsReady; // Returns true or false
			Obj.volname = oDrive.VolumeName;
			Obj.totalsize = (oDrive.TotalSize/1024); // In Kbytes
			Obj.availspace = (oDrive.AvailableSpace/1024); // In Kbytes
			Obj.freespace = (oDrive.FreeSpace/1024); // In Kbytes
			Obj.serial = oDrive.SerialNumber;
			try{Obj.share = oDrive.ShareName}
			catch(ee){Obj.share = null}
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return Obj;
	}
}

function js_io_drivesdelete(iOpt,sDriveStream){
	try{
		var bResult = false;
		if(iOpt == 1){ //  Deletes all drives
			var sCmd = "%comspec% /c net use * /del /y";
			if(oWsh.Run(sCmd,oReg.hide,true) == 0) bResult = true;
		}
		else if(iOpt == 2){ // deletes as many as you wish
			for(var sCmd, sDrive, i = 1; i < arguments.length; i++){
				if(oFso.DriveExists(sDrive = arguments[i])){
					try{
						oWno.RemoveNetworkDrive(sDrive,true);
						bResult = true;
					}
					catch(ee){
						js.err.description = ee.description;
						bResult = false;
					}
				}
			}
		}
		else if(iOpt == 3){ // Del IPC$
			sCmd = "%comspec% /c net use \\\\" + sDriveStream + "\\IPC$ /del /y";
			if(oWsh.Run(sCmd,oReg.hide,true) == 0) bResult = true;
		}		
		return bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_io_drivefloppy(iOpt,sFile,sMessage,iLimit,iIter,sTitle){
	try{
		var oFloppy = oFso.GetDrive("A:");
		sTitle = sTitle ? sTitle : "Floppy";
		iIter = js_str_isnumber(iIter) ? iIter : 1;
		iLimit = js_str_isnumber(iLimit) ? iLimit : 1200; // 20 minutes
		if(iOpt == 1){ // Insert floppy
			sMessage = sMessage ? sMessage : "Insert floppy into drive";
			while(!oFloppy.IsReady || !oFso.FileExists(sFile)){
				if(iIter <= 4){
					oWsh.Popup(iIter++ + "(4) - " + sMessage,iLimit,sTitle,48+0);
					continue;
				}
				else break;
			}
		}
		else if(iOpt == 2){ // Remove Floppy
			sMessage = sMessage ? sMessage : "Remove floppy from drive";
			while(oFloppy.IsReady){
				if(iIter <= 4) oWsh.Popup(iIter++ + "(4) - " + sMessage,iLimit,sTitle,48+0);
				else break;
			}
		}
		else return false;
		return (iIter <= 4);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_drive_firstavailable(){
	try{
		var oDrives = new Enumerator(oFso.Drives);  //Create Enumerator object.
		oDrives.moveFirst();                   //Move to first drive.
		s = "";                          //Initialize s.
		do{
			var x = oDrives.item();                 //Test for existence of drive.
			if (x.IsReady){	
				s = x.DriveLetter + ":";   //Assign 1st drive letter to s.
				break;
			}
			else if (oDrives.atEnd()){
				s = "No drives are available";
            			break;
         		}
      			oDrives.moveNext();                 //Move to the next drive.
		}
   		while (!oDrives.atEnd());              //Do while not at collection end.
   		return(s);                       //Return list of available drives.
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/*

function ShowDriveList(){
   var fso, s, n, e, x;
   fso = new ActiveXObject("Scripting.FileSystemObject");
   e = new Enumerator(fso.Drives);
   s = "";
   for (; !e.atEnd(); e.moveNext())
   {
      x = e.item();
      s = s + x.DriveLetter;
      s += " - ";
      if (x.DriveType == 3)
         n = x.ShareName;
      else if (x.IsReady)
         n = x.VolumeName;
      else
         n = "[Drive not ready]";
      s +=   n + "<br>";
   }
   return(s);
}

*/

