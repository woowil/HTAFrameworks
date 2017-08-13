// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.

//**Start Encode**


////////////////////////////////////////////////////////////////////////////////////////////////
/////// SHELL APPLICATION FUNCTIONS (http://userpages.umbc.edu/~kbradl1/wsz/ref/ShellAppref.html)
/////// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/shell_basics/shell_basics_programming/objectmap.asp
////////////////////////////////////////////////////////////////////////////////////////////////

/*

BrowseForFolder(Hwnd, strTitle, constFlags, strRootDir or constDirectory) returns shellfolder object
displays the "Browse for Folder" dialog and starts at the Root directory;
Hwnd is the name of a form or 0 if no form exists; path of chosen directory is object.Items().Item.Path

constFlags:
&H0001    BIF_returnonlyfsdirs
&H0002    BIF_dontgobelowdomain
&H0004    BIF_statustext 
&H0008 BIF_returnfsancestors
&H0010 BIF_editbox &H0020 BIF_validate 
&H1000 BIF_browseforcomputer
&H2000  BIF_browseforprinter
&H4000 BIF_browseincludefiles 

constDirectory:
15 ssfBITBUCKET  
33 ssfCOOKIES
00 ssfDESKTOP 
16 ssfDESKTOPDIRECTORY
17 ssfDRIVES
06 ssfFAVORITES 
18 ssfNETWORK
4  ssfPRINTERS
38 ssfPROGRAMFILES 

*/

__H.include("HUI@IO.js")

__H.register(__H.IO,"Shell","Shell Application",function Shell(){
	var shell	
	var oo = {
		DESKTOP			: 0x0, //Desktop (including system items)
		PROGRAMS			: 0x2, //Programs section of Start Menu
		CONTROLS			: 0x3, //Control Panel (no path)
		PRINTERS			: 0x4, //Printers (no path)
		PERSONAL			: 0x5, //My Documents
		FAVORITES			: 0x6, //IE Favorites
		STARTUP			: 0x7, //Startup
		RECENT			: 0x8, //Recent
		SENDTO			: 0x9, //SendTo
		BITBUCKET			: 0x10, //Recycle Bin (no path)
		STARTMENU			: 0x11, //Start Menu
		DESKTOPDIRECTORY	: 0x16, //Desktop directory (no system items)
		DRIVES			: 0x17, //My Computer
		NETWORK			: 0x18, //Network Neighborhood
		NETHOOD			: 0x19,
		FONTS				: 0x20,
		TEMPLATES			: 0x21, //The ShellNew directory
		COMMONSTARTMENU	: 0x22, //All users start menu
		COMMONPROGRAMS	: 0x23,
		COMMONSTARTUP		: 0x24,
		COMMONDESKTOPDIR	: 0x25,
		APPDATA			: 0x26, //Application Data directory
		PRINTHOOD			: 0x27,
		LOCALAPPDATA		: 0x28,
		ALTSTARTUP		: 0x29,
		COMMONALTSTARTUP	: 0x30,
		COMMONFAVORITES	: 0x31, //All Users Favorites
		INTERNETCACHE		: 0x32, //Temporary Internet Files
		COOKIES			: 0x33, //Cookies
		HISTORY			: 0x34, //History (no path)
		COMMONAPPDATA		: 0x35,
		WINDOWS			: 0x36,
		SYSTEM			: 0x37,
		PROGRAMFILES		: 0x38,
		MYPICTURES		: 0x39,
		PROFILE			: 0x40,
		SYSTEMx86			: 0x41,
		PROGRAMFILESx86	: 0x48,
		
		BIF_SHOWALLOBJECTS	: 0x0001, //Blocks display of non-file items
		BIF_SHOWEXTENSIONS	: 0x0002, //Changes inside prompt to "Select a File"
		BIF_SHOWCOMPCOLOR	: 0x0008,
		BIF_SHOWSYSFILES	: 0x0032,
		BIF_WIN95CLASSIC	: 0x0064,
		BIF_DOUBLECLICKINWEBVIEW : 0x0128,
		BIF_DESKTOPHTML	: 0x0512,
		BIF_EDITBOX : 0x0010,
		BIF_VALIDATE : 0x0020,
		BIF_NONEWFOLDER : 0x0200,
		BIF_BROWSEFORCOMPUTER : 0x1000,
		BIF_BROWSEFORPRINTER : 0x2000,
		BIF_BROWSEINCLUDEFILES : 0x4000,
		BIF_SHOWFILES	: 0x16384 //Displays ALL files	
	}
	
	var initialize = function initialize(){
		try{ // Works in Win98 and above. Works only in WinNT if IE has activeX (latest shell.dll) 			
			shell = new ActiveXObject("Shell.Application")
			shell.NameSpace("C:\\")
		}
		catch(ee){		
			__HLog.error(ee,this)
			throw new Error(__HLog.errorCode("error"),"Unable to load Shell.Application")
		}
	}
	initialize()
	
	this.browseFolder = function browseFolder(sPath,sTitle,hWindow,hView){
		try{
			var sFolder = false;
			sPath = oFso.FolderExists(sPath) ? sPath : oo.DESKTOP; // Buggy... must check this
			hWindow = hWindow ? hWindow : 0x0;
			hView = hView ? hView : 0x0;
			sTitle = sTitle ? sTitle : "Browsing folder from: " + sPath
			
			var oFolder = shell.BrowseForFolder(hWindow,sTitle,hView,sPath);
			if(oFolder != null){
				//sFolder = oFolder.ParentFolder.ParseName(oFolder.Title).Path;
				sFolder = (oFolder != "Desktop") ? ( oFolder.Items().Item() ).Path : oWsh.SpecialFolders(oFolder);
	 		}
			return sFolder;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.browseFile = function browseFile(sTitle,hWindow,hView,sPath){
		try{
			var sFile = false, oFile;
			try{ // XP Style http://blogs.msdn.com/gstemp/archive/2004/02/17/74868.aspx__H
				var oUA = new ActiveXObject("UserAccounts.CommonDialog")
				if(oFile = oUA.ShowOpen()){
					sFile = oUA.FileName;
				}
				return sFile
			}
			catch(ee){}
	    	hWindow = hWindow ? hWindow : 0x0
			sTitle = typeof(sTitle) == "string" ? sTitle : "Choose File.."
			hView = hView ? hView : 0x4031;
			sPath = sPath ? sPath : 0x0011;
			var oFile = shell.BrowseForFolder(hWindow,sTitle,hView,sPath);
			if(oFile != null){
		    	sFile = oFile.ParentFolder.ParseName(oFile.Title).Path;
	 		}
			return sFile;
		}
		catch(ee){
			//Invalid procedure call or argument
			__HLog.error(ee,this)
			return false;
		}
	}

	this.open = function open(){
		try{
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
			    if(oFso.FolderExists(arguments[i]) || oFso.FileExists(arguments[i])){		    	
					shell.open(arguments[i]);
			    }
			}
	    }
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setWindow = function setWindow(sOpt){
		try{
		   if(typeof(shell) != "object") return false
		    if(sOpt == "minimizeall") shell.MinimizeAll();
		    else if(sOpt == "restore") shell.UndoMinimizeAll();
		    else if(sOpt == "minimize") shell.Minimize();
			else if(sOpt == "refresh"){
				shell.MinimizeAll();
				shell.UndoMinimizeAll();
			}
	    }
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

})

var __HShell = new __H.IO.Shell()