// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-gui-shell.js
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
	var oShX = new ActiveXObject("Shell.Application");
	oShX.NameSpace("C:\\")
}
catch(ee){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
	var oShX = new ActiveXObject("Shell.Application");
	oShX.NameSpace("C:\\")
}

var oSh = new js_shl_object();

////////////////////////////////////////////////////////////////////////////////////////////////
/////// SHELL APPLICATION FUNCTIONS (http://userpages.umbc.edu/~kbradl1/wsz/ref/ShellAppref.html)
/////// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/shell_basics/shell_basics_programming/objectmap.asp
////////////////////////////////////////////////////////////////////////////////////////////////

/*

BrowseForFolder(Hwnd, strTitle, constFlags, strRootDir or constDirectory) returns shellfolder object
displays the "Browse for Folder" dialog and starts at the Root directory; Hwnd is the name of a form or 0 if no form exists; path of chosen directory is object.Items().Item.Path
constFlags:
&H0001    BIF_returnonlyfsdirs  &H0002    BIF_dontgobelowdomain     &H0004    BIF_statustext 
&H0008 BIF_returnfsancestors &H0010 BIF_editbox &H0020 BIF_validate 
&H1000 BIF_browseforcomputer     &H2000  BIF_browseforprinter &H4000 BIF_browseincludefiles 
constDirectory:
15    ssfBITBUCKET  33    ssfCOOKIES     0    ssfDESKTOP 
16 ssfDESKTOPDIRECTORY 17 ssfDRIVES 6 ssfFAVORITES 
18 ssfNETWORK     4  ssfPRINTERS 38 ssfPROGRAMFILES 


*/

function js_shl_object(){
	// Uses the "Shell.Application" Object. Works in Win98 and above. Works only in WinNT if IE has activeX (latest shell.dll) 
	//Directory constants
	this.DESKTOP			= 0x0 //Desktop (including system items)
	this.PROGRAMS			= 0x2 //Programs section of Start Menu
	this.CONTROLS			= 0x3 //Control Panel (no path)
	this.PRINTERS			= 0x4 //Printers (no path)
	this.PERSONAL			= 0x5 //My Documents
	this.FAVORITES			= 0x6 //IE Favorites
	this.STARTUP			= 0x7 //Startup
	this.RECENT			= 0x8 //Recent
	this.SENDTO			= 0x9 //SendTo
	this.BITBUCKET			= 0x10 //Recycle Bin (no path)
	this.STARTMENU			= 0x11 //Start Menu
	this.DESKTOPDIRECTORY	= 0x16 //Desktop directory (no system items)
	this.DRIVES			= 0x17 //My Computer
	this.NETWORK			= 0x18 //Network Neighborhood
	this.NETHOOD			= 0x19
	this.FONTS				= 0x20
	this.TEMPLATES			= 0x21 //The ShellNew directory
	this.COMMONSTARTMENU	= 0x22 //All users start menu
	this.COMMONPROGRAMS	= 0x23
	this.COMMONSTARTUP		= 0x24
	this.COMMONDESKTOPDIR	= 0x25
	this.APPDATA			= 0x26 //Application Data directory
	this.PRINTHOOD			= 0x27
	this.LOCALAPPDATA		= 0x28
	this.ALTSTARTUP		= 0x29
	this.COMMONALTSTARTUP	= 0x30
	this.COMMONFAVORITES	= 0x31 //All Users Favorites
	this.INTERNETCACHE		= 0x32 //Temporary Internet Files
	this.COOKIES			= 0x33 //Cookies
	this.HISTORY			= 0x34 //History (no path)
	this.COMMONAPPDATA		= 0x35
	this.WINDOWS			= 0x36
	this.SYSTEM			= 0x37
	this.PROGRAMFILES		= 0x38
	this.MYPICTURES		= 0x39
	this.PROFILE			= 0x40
	this.SYSTEMx86			= 0x41
	this.PROGRAMFILESx86	= 0x48
	
	this.BIF_SHOWALLOBJECTS	= 0x0001 //Blocks display of non-file items
	this.BIF_SHOWEXTENSIONS	= 0x0002 //Changes inside prompt to "Select a File"
	this.BIF_SHOWCOMPCOLOR	= 0x0008
	this.BIF_SHOWSYSFILES	= 0x0032
	this.BIF_WIN95CLASSIC	= 0x0064
	this.BIF_DOUBLECLICKINWEBVIEW = 0x0128
	this.BIF_DESKTOPHTML	= 0x0512
	this.BIF_EDITBOX = 0x0010
	this.BIF_VALIDATE = 0x0020
	this.BIF_NONEWFOLDER = 0x0200
	this.BIF_BROWSEFORCOMPUTER = 0x1000
	this.BIF_BROWSEFORPRINTER = 0x2000
	this.BIF_BROWSEINCLUDEFILES = 0x4000
	this.BIF_SHOWFILES	= 0x16384 //Displays ALL files	
}

function js_shl_browsefolder(sTitle,hWindow,hView,sPath){
	try{
		if(!oShX) return false
		var sFolder = false;
		hWindow = hWindow ? hWindow : 0x0
		sTitle = sTitle ? sTitle : "Choose Folder.."    
		hView = hView ? hView : 0x10;
		sPath = sPath ? sPath : 0x11;
		var oFolder = oShX.BrowseForFolder(hWindow,sTitle,hView, sPath);
		if(oFolder != null){
			//sFolder = oFolder.ParentFolder.ParseName(oFolder.Title).Path;
			sFolder = (oFolder != "Desktop") ? ( oFolder.Items().Item() ).Path : oWsh.SpecialFolders(oFolder);
 		}
		return sFolder;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_shl_browsefile(sTitle,hWindow,hView,sPath){
	try{
		if(!oShX) return false
		var sFile = false, oFile;
		try{ // XP Style http://blogs.msdn.com/gstemp/archive/2004/02/17/74868.aspx
			var oUA = new ActiveXObject("UserAccounts.CommonDialog")
			if(oFile = oUA.ShowOpen()){
				sFile = oUA.FileName;
			}
			return sFile
		}
		catch(ee){}
    	hWindow = hWindow ? hWindow : 0x0
		sTitle = sTitle ? sTitle : "Choose File.."
		hView = hView ? hView : 0x4031;
		sPath = sPath ? sPath : 0x0011;
		var oFile = oShX.BrowseForFolder(hWindow,sTitle,hView,sPath);
		if(oFile != null){
	    	sFile = oFile.ParentFolder.ParseName(oFile.Title).Path;
 		}
		return sFile;
	}
	catch(e){
		//Invalid procedure call or argument
		js_log_error(2,e);
		return false;
	}
}

function js_shl_open(sFolderFile){
	try{
	    if(!oShX) return false
	    if(oFso.FolderExists(sFolderFile) || oFso.FileExists(sFolderFile)){
	    	oShX.Open(sFolderFile);
	    }
    }
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_shl_window(iOpt){
	try{
	    if(!oShX) return false
	    if(iOpt == 1) oShX.MinimizeAll();
	    else if(iOpt == 2) oShX.UndoMinimizeAll();
	    else if(iOpt == 3) oShX.Minimize();
    }
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

