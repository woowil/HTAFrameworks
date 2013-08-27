// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.register(__H,"IO","drive management, input/output",function IO(){
	
	// File modes
	this.read = 1
	this.write = 2
	this.append = 8
	this.TristateTrue = -1// Open the file as Unicode
	this.TristateFalse = 0// Open the file as ASCII
	this.TristateUseDefault = -2// Open the file using theSystem default

	// Window Property for oWsh.Run()
	this.hide = 0// Hides the window and activates another window.
	this.show = 1, // Activates and displays a window. If the window is minimized or maximizedtheSystem restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
	this.actmin = 2// Activates the window and displays it as a minimized window.
	this.actmax = 3// Activates the window and displays it as a maximized window.
	this.showmin = 7// Displays the window as a minimized window. The active window remains active.
	
	// Windows Directory
	this.windir = oFso.GetSpecialFolder(0)
	this.system32 = oFso.GetSpecialFolder(1)
	this.temp = oFso.GetSpecialFolder(2)
	// https://windowsxp.mvps.org/usersshellfolders.htm
	this.appdata_all = oWsh.ExpandEnvironmentStrings("%allusersprofile%") + "\\Application Data"
	this.appdata = oWsh.ExpandEnvironmentStrings("%userprofile%") + "\\Application Data"
	
	this.w32 = oFso.GetSpecialFolder(1) + "\\spool\\drivers\\w32x86"
	this.w32_1 = oFso.GetSpecialFolder(1) + "\\1"
	this.w32_2 = oFso.GetSpecialFolder(1) + "\\2"
	this.w32_3 = oFso.GetSpecialFolder(1) + "\\3"
	
	this.getDriveInfo = function getDriveInfo(sDrive){
		try{
			var o = {}
			var d = oFso.GetDrive(oFso.GetDriveName(sDrive))
			switch(d.DriveType){
				case 0: o.Type = "unknown"; break;
				case 1: o.Type = "removable"; break;
				case 2: o.Type = "fixed"; break; // Both regular drive and substitute drives
				case 3: o.Type = "network"; break;
				case 4: o.Type = "cdrom"; break;
				case 5: o.Type = "ramdisk"; break;
				default : o.Type = null; break;
			}
			if(o.Type != null){
				o.DriveLetter = d.DriveLetter;
				o.IsReady = d.IsReady; // Returns true or false
				o.VolumeName = d.VolumeName;
				o.TotalSize = (d.TotalSize/1024); // In Kbytes
				o.AvailableSpace = (d.AvailableSpace/1024); // In Kbytes
				o.FreeSpace = (d.FreeSpace/1024); // In Kbytes
				o.SerialNumber = d.SerialNumber;
				try{
					o.ShareName = d.ShareName
				}
				catch(eee){
					o.ShareName = null
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getDrives = function getDrives(){
		try{
			var a = []			
			for(var oEnum = new Enumerator(oFso.Drives); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(), v, t, bReady = false
				switch(oItem.DriveType){
					case 0: t = "Unknown Disk"; break;
					case 1: t = "Removable Drive"; break; // Including floppy
					case 2: t = "Local Disk"; break; // Both regular drive and substitute drives
					case 3: t = "Network Drive"; break;
					case 4: t = "CD-ROM Drive"; break;
					case 5: t = "RAM Drive"; break;
					default : t = "Unknown Disk"; break;
				}
				
				if(oItem.IsReady) v = oItem.VolumeName, bReady = true
				else if(t == 3) v = oItem.ShareName
				else v = "Not Ready"
				
				var o = {
					DriveLetter	: oItem.DriveLetter + ":",
					DriveType	: oItem.DriveType,
					DriveTypeName : t,
					VolumeName	: v,
					TotalSize	: (bReady ? __HSys.bytes(oItem.TotalSize) : 0),
					FileSystem	: (bReady ? oItem.FileSystem : "Unknown"),
					AvailableSpace : (bReady ? __HSys.bytes(oItem.AvailableSpace): 0),
					SerialNumber : (bReady ? oItem.SerialNumber : "Unknown"),
					isReady : bReady
				}
				o.Text = o.DriveLetter + ", " + o.VolumeName + ", " + o.FileSystem + ", " + o.TotalSize + ", " + o.AvailableSpace + ", " + o.DriveType + ", " + o.SerialNumber
				a.push(o)
			 }
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return a
	}
	
	this.deleteAllDrives = function deleteAllDrives(){
		try{
			return (oWsh.Run("%comspec% /c net use * /del /y",__HIO.hide,true) == 0)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.deleteDrive = function deleteDrive(sDrive){
		try{
			if(oFso.DriveExists(sDrive)){
				var d = oFso.GetDrive(oFso.GetDriveName(sDrive))
				if(d.DriveType == 3){ // Network
					oWno.RemoveNetworkDrive(sDrive,true)
					return true
				}
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	//// FLOPPY
	
	this.floppyInsert = function floppyInsert(sMessage){
		try{
			var oFloppy = oFso.GetDrive("A:");
			sMessage = sMessage ? sMessage : "Insert floppy from drive";
			for(var i = 1; oFloppy.IsReady; i++){
				if(i <= 4){
					oWsh.Popup(i++ + "(4) - " + sMessage,1200,"Floppy Insert",48+0);
					continue;
				}
				else break;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.floppyRemove = function floppyRemove(sMessage){
		try{
			var oFloppy = (oFso.GetDrive("A:") || oFso.GetDrive("B:"));
			sMessage = sMessage ? sMessage : "Remove floppy into drive";
			for(var i = 1; oFloppy.IsReady; i++){
				if(i <= 4){
					oWsh.Popup(i + "(4) - " + sMessage,1200,"Floppy Remove",48+0);
					continue;
				}
				else break;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
})


var __HIO = new __H.IO()