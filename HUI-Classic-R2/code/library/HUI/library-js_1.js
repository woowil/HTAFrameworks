// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:	  library-js_1.js
Purpose:  Development script (main Library)
Author:   Woody Wilson
Created:  2002-07-08
Version:
Dependency:  library-js_1.js

Description:
JScript library scripting functions used for WSH files or HTA Applications.

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
		var oXml = new ActiveXObject("Microsoft.XMLDOM");
		oXml.async = false;
	}
}
catch(ee){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
	var oXml = new ActiveXObject("Microsoft.XMLDOM");
	oXml.async = false;
}

var LIB_NAME	= "Main Library";
var LIB_VERSION = "5.0";
var LIB_FILE	= oFso.GetAbsolutePathName("library-js_1.js")

var js = new js_object();
var oReg = new js_reg_object();

function js_object(){
	try{
		this.iswsscript = true;
		WScript.Sleep(1);
	}
	catch(e1){
		this.iswsscript = false;
		try{
			var test = new Array()
			test.push("nada")
		}
		catch(e2){
			js_log_error(1,"PROGRAM FAILURE!!\n\nYou are probably using Internet Explore 5.0 or lower.\nThe script files reguires Internet Explore 5.5 or higher.")
			try{window.close()}
			catch(e3){WScript.Quit(-1)}
		}
	}
	try{
		this.ishta = true
		var t = document.body.className
	}
	catch(e3){
		this.ishta = false
	}
	this.version = new js_objectversion(LIB_NAME,LIB_FILE,LIB_VERSION)
	this.resource = this.computer = oWno.ComputerName;

	try{
		this.shellx = new ActiveXObject("Shell.Application");
	}
	catch(e1){ this.shellx = false;}
	this.autoitx = false; // The ActiveX Control of AutoItX3.dll
	this.regobjx = false; // The ActiveX Control of RegObj.dll
	this.jssysx = false; // The ActiveX Control of JSSys3.Ops
	this.adssecurityx = false  // The ActiveX Control of ADsSecurity.dll
	this.logfile = false // log file
	this.log_res_isactive = log_err_isactive = false
	this.log_maxwait = 15000
	this.caller_max = 10
	this.disclaimer = "This sample code is provided AS IS WITHOUT " +
						"WARRANTY OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY " +
						"OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code " +
						"is to be undertaken entirely at your risk, and the " +
						"results that may be obtained from it are dependent on the user. " +
						"Please note to fully back up files and system(s) on a regular " +
						"basis. A failure to do so can result in loss of data or damage to systems. "
	this.log = new Object()
		this.log.scroll = false // a function to scroll textarea or DIV
	this.err = new Object(); // Error object handled by js_log_error()
		this.err.name = null;
		this.err.func = null;
		this.err.number = null;
		this.err.description = null;
		this.err.textarea_name = false
		this.err.textarea_result = this.err.textarea_error = false // Used in WSH and HTA Applications
		this.err.divtable_error = false
		this.err.logfile = false; // This object is used for WSH or HTA applications
	this.fls = new Object();
		this.fls.sleep = oFso.GetSpecialFolder(2) + "\\js_tme_sleep.js";
		this.fls.SUBLIMIT = 1000000;
	this.pro = new Object()
		this.pro.stopit = false
	this.time = new Object(); // all in milli seconds
		this.time.mil1 = 1;
		this.time.mil2 = 2;
		this.time.mil3 = 3;
		this.time.mil10 = 10;
		this.time.mil50 = 50;
		this.time.mil100 = 100;
		this.time.mil200 = 200;
		this.time.mil500 = 500;
		this.time.sec1 = 1000;
		this.time.sec2 = 2000;
		this.time.sec3 = 3000;
		this.time.sec4 = 4000;
		this.time.sec5 = 5000;
		this.time.sec10 = 10000;
		this.time.sec15 = 15000;
		this.time.sec30 = 30000;
		this.time.sec45 = 45000;
		this.time.sec60 = this.time.min1 = 60000;
		this.time.sec90 = 90000;
		this.time.min2 = 120000;
		this.time.min3 = 180000;
		this.time.min5 = 300000;
		this.time.min15 = 900000;
		this.time.min30 = 1800000;

	this.ore = new Object();
		this.ore.ipaddress = new RegExp("([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})","ig"); // IP-address

	// Jscript VBarray
	this.VBArray = function(args){
		var d = new ActiveXObject("Scripting.Dictionary")
		for(var i = 0, len = arguments.length; i < len; i++){
			d.Add("d" + i,arguments[i])
		}
		return d.Items()
	}

	this.isString = function(args){
		for(var s, i = 0, len = arguments.length; i < len; i++){
			if(typeof(s = arguments[i]) == "string" && (s.replace(/[ \t]*(.*)[ \t]*$/g,"$1")).length > 0) continue
			else return false
		}
		return true
	}
}

function js_objectversion(sName,sFile,sVersion){
	this.Name = sName
	this.File = sFile
	this.Version = sVersion
}

function js_reg_object(){
	try{
		this.HKCR = 0x80000000;
		this.HKCU = 0x80000001;
		this.HKLM = 0x80000002;
		this.HKU  = 0x80000003;
		this.HKCC = 0x80000005;

		this.REG_SZ = 1;
		this.REG_EXPAND_SZ = 2;
		this.REG_BINARY = 3;
		this.REG_DWORD = 4;
		this.REG_DWORD_BIG_ENDIAN = 5;
		this.REG_LINK = 6;
		this.REG_MULTI_SZ = 7;
		this.REG_RESOURCE_LIST = 8;
		this.REG_FULL_RESOURCE_DESCRIPTOR = 9;
		this.REG_RESOURCE_REQUIREMENTS_LIST = 10;
		this.REG_QWORD = 11;

		this.KEY_QUERY_VALUE = 0x1;
		this.KEY_SET_VALUE = 0x2;
		this.KEY_CREATE_SUB_KEY = 0x4;
		this.KEY_ENUMERATE_SUB_KEYS = 0x8;
		this.KEY_NOTIFY = 0x10;
		this.KEY_CREATE_LINK = 0x20;
		this.DELETE = 0x10000;
		this.READ_CONTROL = 0x20000;
		this.WRITE_DAC = 0x40000;
		this.WRITE_OWNER = 0x80000;

		this.hkcr = "HKEY_CURRENT_ROOT";
		this.hkcu = "HKEY_CURRENT_USER";
		this.hklm = "HKEY_LOCAL_MACHINE";
		this.hku  = "HKEY_USERS";
		this.hkcc = "HKEY_CURRENT_CONFIG";

		this.con = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control";
		this.ses = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\Session Manager";
		this.env = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment";
		this.env2 = "SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment";
		this.uenv = this.hklu + "\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment";
		this.mem = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management";
		this.win = this.hklm + "\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
		this.win2 = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon";
		this.cur = this.hklm + "\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion";
		this.cur2 = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion";
		this.srv = this.hklm + "\\SYSTEM\\CurrentControlSet\\Services"
		this.srv_net = this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters" // Act as a NT server "NeutralizeNT4Emulator"=dword:00000001 (must restart netlogon)
		this.run = this.hklm + "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
		this.once = this.hklm + "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce";
		this.prn = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\Print\\Printers";
		this.prn_port = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports";
		this.prn_drv3 = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\Windows NT x86\\Drivers\\Version-3";
		this.prn_ins = this.hklm + "\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Print\\Providers\\LanMan Print Services\\Servers"; // \server.dnbnor.net\Printers
		this.pro = this.hklm + "\\SOFTWARE\\Classes\\Installer\\Products"
		this.drv_sign1 = this.hklm + "\\Software\\Policies\\Microsoft\\Windows NT\\Driver Signing"; // Windows 2000/XP - http://www.winguides.com/registry/display.php/1298
		this.drv_sign2 = this.hklm + "\\Software\\Microsoft\\Driver Signing"; // Policy=(Ignore=00, Warn=01, Block=02) Windows 2003
		this.lan_par = this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\lanmanserver\\parameters"
		this.tcpip = this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" //\\{605DA0E2-2EB2-4ECA-89F0-4EDDAD382641}\\NameServers
		this.tcpip2 = "SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters" //\\{605DA0E2-2EB2-4ECA-89F0-4EDDAD382641}\\NameServers
		this.tcpip_int = this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Interfaces" //\\{605DA0E2-2EB2-4ECA-89F0-4EDDAD382641}\\NameServers
		this.lan_srv =  this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\lanmanserver\\parameters" // srvcomment
		this.nbt_int =  this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters\\Interfaces" // srvcomment
		//IE
		this.ie_cu =  this.hkcu + "\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings"
		this.ie_lm =  this.hklm + "\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings"
		// SNA
		this.sna_par = this.hklm + "\\System\\CurrentControlSet\\Services\\SNABase\\Parameters"
		this.sna_netsoft = this.hklm + "\\SOFTWARE\\NetSoft\\NS/Elite\\CurrentVersion\\Workspaces\\Mainframe Workspace\\Links\\TN3270\\NS/Router 3270"; // Data (BINARY)
		// http://support.microsoft.com/default.aspx?scid=kb;en-us;Q175500
		this.ie_styles = this.hkcu + "\\Software\\Microsoft\\Internet Explorer\\Styles" // MaxScriptStatements
		// NT Computer name
		this.comp1 = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\ComputerName\\ComputerName"; // Computername
		this.comp2 = this.hklm + "\\SYSTEM\\CurrentControlSet\\Control\\ComputerName\\ActiveComputerName";
		this.comp3 = this.hklm + "\\CLONE\\CLONE\\Services\\Tcpip\\Parameters"; // Key: Hostname
		this.comp4 = this.hklm + "\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters"; // Key: Hostname
		this.oem = this.hklm + "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate\\OemInfo";

		// Specific Troll settings
		this.ino6 = this.hklm + "\\SOFTWARE\\ComputerAssociates\\InoculateIT\\6.0\\NameCli"; // ServerList
		this.acti = this.hklm + "\\SOFTWARE\\Microsoft\\Active Setup\\Installinfo"; // UNCDownloadDir
		this.jiti = this.hklm + "\\SOFTWARE\\Microsoft\\Active Setup\\JITinfo\\5"; // UNCDownloadDir
		this.odbc = this.hklm + "\\SOFTWARE\\Microsoft\\MSSQLServer\\Client\\ConnectTo"; // localserver
		this.tn3270 = this.hklm + "\\SOFTWARE\\NetSoft\\NS/Elite\\CurrentVersion\\Workspaces\\Mainframe Workspace\\Links\\TN3270\\NS/Router 3270"; // Data (BINARY)

		// File modes
		this.read = 1;
		this.write = 2;
		this.append = 8;
		this.TristateTrue = -1; // Open the file as Unicode
		this.TristateFalse = 0; // Open the file as ASCII
		this.TristateUseDefault = -2; // Open the file using the system default

		// RegObj
		this.key = new ActiveXObject("Scripting.Dictionary");
		this.key.add(0,"REG_NONE"); // No value type
		this.key.add(1,"REG_SZ"); // String, A sequence of characters representing human readable text. Unicode null terminated string.
		this.key.add(2,"REG_EXPAND_SZ"); // String, An expandable data string, which is text that contains a variable to be replaced when called by an application (ex : %windir%\system\wsock32.dll ). Unicode null terminated string.
		this.key.add(3,"REG_BINARY"); // VBArray Raw binary data. Most hardware component information is stored as binary data, and can be displayed in hexadecimal format
		this.key.add(4,"REG_DWORD"); // = REG_DWORD_LITTLE_ENDIAN, Number, 32 bits number
		this.key.add(5,"REG_DWORD_BIG_ENDIAN"); // 32 bits number but in big endian format
		this.key.add(6,"REG_LINK"); // A symbolic link (unicode)
		this.key.add(7,"REG_MULTI_SZ"); // VBArray, A multiple string. Values that contain lists or multiple values in human readable text are usually this type (unicode). Entries are separated by NULL characters.
		this.key.add(8,"REG_RESOURCE_LIST"); // Resource list in the resource map
		this.key.add(9,"REG_FULL_RESOURCE_DESCRIPTOR"); // Resource list in the hardware description
		this.key.add(10,"REG_RESSOURCE_REQUIREMENT_MAP"); // Resource list in the hardware description

		// Window Property for oWsh.Run()
		this.hide = 0; // Hides the window and activates another window.
		this.show = 1; // Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
		this.actmin = 2 // Activates the window and displays it as a minimized window.
		this.actmax = 3; // Activates the window and displays it as a maximized window.
		this.showmin = 7; // Displays the window as a minimized window. The active window remains active.

		// Windows Directory
		this.windir = oFso.GetSpecialFolder(0);
		this.system32 = oFso.GetSpecialFolder(1);
		this.temp = oFso.GetSpecialFolder(2);
		// https://windowsxp.mvps.org/usersshellfolders.htm
		this.appdata_all = oWsh.ExpandEnvironmentStrings("%allusersprofile%") + "\\Application Data"
		this.appdata = oWsh.ExpandEnvironmentStrings("%userprofile%") + "\\Application Data"
		this.w32 = oFso.GetSpecialFolder(1) + "\\spool\\drivers\\w32x86";
		this.w32_1 = this.w32 + "\\1";
		this.w32_2 = this.w32 + "\\2";
		this.w32_3 = this.w32 + "\\3";
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////
/////// STRING FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////


function js_dict_getkeys(sOpt,oDict,sKeyGet,bIgnoreCase){ // Returns dictionary keys
	try{
		var aKeys = (new VBArray(oDict.Keys())).toArray();
		var oValue = false, sKey;
		sKeyGet = bIgnoreCase ? sKeyGet.toLowerCase() : sKeyGet;
		for(var i = 0, len = aKeys.length; i < len; i++){
			sKey = bIgnoreCase ? aKeys[i].toLowerCase() : aKeys[i];
			sKey = (sOpt == "getkey") ? sKey : oDict(sKey); // iOpt=1: Searches key, iOpt=2: Searches item
			if(sKey == sKeyGet){
				oValue = new Object();
				oValue.key = aKeys[i];
				oValue.item = oDict(aKeys[i]);
				oValue.idx = i;
				break;
			}
		}
		return oValue;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_popup(sMsg,iIdleMilli,sTitle,nType){
	try{
		if(typeof(sTitle) != "string") var sTitle = "Popup";

		iIdleSec = js_str_isnumber(iIdleMilli) ? (iIdleMilli/1000) : 30;
		/* nType = Button Types + Icon Types
		Button Types
		0 Show OK button.
		1 Show OK and Cancel buttons.
		2 Show Abort, Retry, and Ignore buttons.
		3 Show Yes, No, and Cancel buttons.
		4 Show Yes and No buttons.
		5 Show Retry and Cancel buttons.

		Icon Types
		16 Show "Stop Mark" icon.
		32 Show "Question Mark" icon.
		48 Show "Exclamation Mark" icon.
		64 Show "Information Mark" icon.
		*/
		nType = js_str_isnumber(nType) ? nType : 48;

		/* IntButton returns
		1 OK button
		2 Cancel button
		3 Abort button
		4 Retry button
		5 Ignore button
		6 Yes button
		7 No button
		*/
		try{ // Fix for HTA applications, since idle ain't working
			sTitle = sTitle + "-#"
			if(iIdleSec > 0){
				var ID = window.setTimeout("js_app_close('" + sTitle + "',0)",(iIdleSec*1000),"JavaScript")
			}
		}
		catch(ee){
			sTitle = sTitle.replace(/(.+)-#$/g,"$1")
		}
		return oWsh.Popup(sMsg,iIdleSec,sTitle,nType);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_compare(sOpt,sStr1,sStr2){
	try{
		if(!js_str_isdefined(sStr1,sStr2)) return false
		else {
			sStr1 = (sStr1.toString()).toLowerCase()
			sStr2 = (sStr2.toString()).toLowerCase()
			return (sStr1.length == sStr2.length && sStr1 === sStr2); // Identical type and same
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_randomize(lval,hval){
	try{
		if(typeof hval != "number") return hval;
		if(typeof lval != "number") lval = 0;
		return Math.ceil(Math.random()*hval + lval)
	}
	catch(e){
		js_log_error(2,e);
		return 0;
	}
}

function js_str_conv_uniascii(iOpt,UniAscii){
	try{
		if(iOpt == 1){ // If not a number
			if(isNaN(UniAscii)){ // FROM ASCII to unicode
				var sChar = UniAscii.toString();
				return (sChar.charCodeAt(0));
			}
		}
		else if(iOpt == 2){ // From Unicode to ASCII
			if(js_str_isnumber(UniAscii)){ // If a number
				var iNumber = new Number(UniAscii);
				return String.fromCharCode(iNumber);
			}
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_conv_digit(iDecimal){
	try{
		var o = false, iDecimal;
		if(js_str_isnumber(iDecimal = parseInt(iDecimal))){
			o = new Object();
			o.Decimal = iDecimal;
			o.Hex = iDecimal > 0 ? iDecimal.toString(16) : (iDecimal + 0x100000000).toString(16)
			o.Octal = iDecimal.toString(8);
			o.Binary = iDecimal.toString(2);
		}
		return o
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_conv_asciihex(sOpt,sString){
	var sConvert = "";
	if(!js_str_isdefined(sOpt,sString)) return false
	if(sOpt == "ascii"){
		for(var i = 0, len = sString.length; i < len; i++){
			sConvert = sConvert +  "%" + sString.charCodeAt(i).toString(16);
		}
	}
	else if(sOpt == "hex"){
		sConvert = unescape(sString);
	}
	return sConvert;
}

function js_str_hex(iNum){
	try{
		if(!js_str_isnumber(iNum)) return false
		else if(iNum > 0) return iNum.toString(16)
		else return (iNum + 0x100000000).toString(16);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_number(iNumber){
	try{
		return (iNumber <= 9) ? "00" + iNumber : ((iNumber <= 99) ? "0" + iNumber : iNumber);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_capitalize(sString,bDont){
	try{
		if(typeof(sString) != "string") return "";
		else if(!bDont && sString.length < 3) return sString.toUpperCase()
		else if(!bDont && sString.search(/.+-.+/ig) > -1) return sString.substring(0,1).toUpperCase() + sString.substring(1,sString.length)
		
		var aString = sString.split(/[ \t]+/g)
		for(var i = 0, len = aString.length; i < len; i++){
			var C = ((aString[i]).substring(0,1)).toUpperCase()
			var R = new String(((aString[i]).substring(1,(aString[i]).length)))
			R = C.search(/[0-9]/) > -1 ? R : R.toLowerCase()
			if(i == 0) sString = C + R
			else sString = sString + " " + C + R
		}
		return sString
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_removechar(sRemove,sChar){
	try{
		var bResult = "";
		for(var j = 0, sCharTmp, len = sRemove.length; j < len; j++){
			if((sCharTmp = sRemove.substring(j,j+1)) != sChar){
				bResult += sCharTmp;
			}
		}
		return bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_kill(){
	try{
		for(var i = 0, l = arguments.length; i < l; i++){
			//prn(i + ": " + arguments[i].constructor)
			if(typeof arguments[i] == "undefined" || arguments[i] == null) continue
			else if(arguments[i] instanceof Function) continue
			else if(arguments[i] instanceof Array){
				//prn("# del ary " + i)
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
					//prn("# del obj " + i)
					for(var o in arguments[i]) js_str_kill(arguments[i][o])
				}
				//else prn("# del enum")
			}
			else if(typeof arguments[i] == "object"){
				if(typeof arguments[i].RemoveAll == "unknown") arguments[i].RemoveAll()//, prn("# del dict " + i)
				//else prn("# del NOT " + i)
			}
			//else prn("# del som " + i)
			//prn("## del")
			delete arguments[i] // delete: http://users.adelphia.net/~daansweris/js/special_operators.html
		}
		return true;
	}
	catch(e){
		//js_log_error(2,e);
		return false;
	}
}


/**
	Author:	 Woody Wilson
	Function:	js_str_trim()
	Parameters: sString
				sString (string) = A line or string to trim.
	Purpose:	Removes new line a spaces at the end and in the beginning
	Returns:	String
	Exception:  On Error
*/
function js_str_trim(s){
	try{
		var ii
		s = s.replace(/([ \t\n]*)(.*)/g,"$2"); // Removes space at beginning
		if((ii = s.search(/([ \t\n]*$)/m)) > 0) s = s.substring(0,ii);
		return s.replace(/[ \t]+/ig," "); // Replaces tabs, doubble spaces
	}
	catch(e){
		return "";
	}
}

/**
	Author:	 Woody Wilson
	Function:	js_str_blankline()
	Parameters: sLine
				sLine (string) = Line to isSearch
	Purpose:	Matches a blank line
	Returns:	Boolean
	Exception:  On Error
*/
function js_str_blankline(sLine){
	return typeof(sLine) == "string" && (sLine.search(/^\[ \t]*$/g) > -1 || sLine.length == 0);
}

function js_str_isalphanum(sString){
	try{
		if(typeof(sLine) != "string") return false;
		var oRe = /[a-z0-9]+/ig, sChar;
		for(var i = 0, len = sString.length; i < len; i++){
			var sChar = sString.substring(i,i+1);
			if(sChar.search(oRe) == -1){
				return false;
			}
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_isnumber(){
	try{
		for(var i = 0, len = arguments.length; i < len; i++){
			if(typeof(arguments[i]) == "object" || arguments[i] == null) return false // Note! null is an object
			else if(isNaN(arguments[i])) return false
		}
		return (i > 0);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
	Author:	 Woody Wilson
	Function:	js_str_isdefined()
	Parameters: sStream
				sStream (string,object,boolean,number) = One or more arguments
	Purpose:	Check if all arguments are valid
	Returns:	Boolean
	Exception:  On Error
*/
function js_str_isdefined(arg){
	try{
		for(var s, i = 0, len = arguments.length; i < len; i++){
			if(typeof(s = arguments[i]) == "object" && s != null) continue // Note! null is an object
			else if((s) == null || s == false || s == undefined) return false
			else if(typeof(s) == "string" && js_str_trim(s) != "") continue
		}
		return (i > 0);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_str_replace(sString,sSplit,sReplace){
	try{
		var aString = sString.split(sSplit);
		for(var j = 0, sString = "", len = aString.length; j < len; j++){
			sString += aString[j] + (j+1 == aString.length ? "" : sReplace);
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return sString;
	}
}


////////////////////////////////////////////////////////////////////////////////////////////////
/////// TIME FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////

function js_tme_stop(ID){
	try{
		if(ID != null){
			clearTimeout(ID);
			clearInterval(ID)
			ID = null
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_tme_sleep(iMilli,bOnce){ // This function is for HTA Applications since it doesn't support WScript.* functions
	try{
		if(typeof(iMilli) != "number") return;
		//if(js.autoitx) js.autoitx.Sleep(iMilli);
		if(js.iswsscript) WScript.Sleep(iMilli);
		else {
			oWsh.Run("%comspec% /c sleep -m " + iMilli + ">nul",oReg.hide,true)
			/*
			if(oFso.FileExists(js.fls.sleep) && js_str_isnumber(iMilli)){
				iMilli = (iMilli < 365) ? 1 : (iMilli-364);
				oWsh.Run("%comspec% /c cscript.exe //nologo \"" + js.fls.sleep + "\" " + iMilli,oReg.hide,true);
			}
			else {
				js_wsh_create(js.fls.sleep,"Sleep","js_tme_sleep()",js_wsh_sleep);
				if(bOnce){
					return js_tme_sleep(iMilli,true);
				}
			}*/
		}
	}
	catch(e){
		js_log_error(2,e);
	}
}


////////////////////////////////////////////////////////////////////////////////////////////////
/////// SYSTEM FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////



function js_sys_ismacaddress(sMacAddress){
	try{
		if(typeof(sMacAddress) != "string") return false;
		else if(sMacAddress.length == 12){ //
			return sMacAddress.search(/[a-z0-9]{12}/ig) > -1;
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_isipaddress(sIPaddress){
	try{
		var bResult = false, aNum;
		if(typeof(sIPaddress) != "string") return false;
		else if(aNum = sIPaddress.match(js.ore.ipaddress)){
			if(sIPaddress.length == aNum[0].length){
				bResult = true;
				for(var i = 1; i < 5; i++){
					if(parseInt(aNum[i]) > 255) bResult = false;
				}
			}
		}
		return bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_localhost(sComputer){
	try{
		if(typeof(sComputer) != "string") return false;
		else if(sComputer.toUpperCase() == oWno.ComputerName.toUpperCase() || sComputer == "127.0.0.1") return true;
		else if(js_sys_isipaddress(sComputer)){
			return (oWsh.Run("%comspec% /c ping -a " + sComputer + " -n 1 -w 50 | find /i \"%computername%\"",oReg.hide,true) == 0);
		}
		/*else if(sIP = js_sys_getipaddress(sComputer)){ // NETBIOS alias
			return js_sys_localhost(sIP);
		}*/
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_info(sOpt,sComputer){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName
		var oSys = new Object();
		if(sOpt == "local"){
			oSys.Computer = oWno.ComputerName;
			oSys.OS = js_reg_envget("SYSTEM","OS");
			oSys.OSversion = js_reg_read(oReg.cur,"CurrentVersion");
			oSys.ServicePack = js_reg_read(oReg.cur,"CSDVersion"); // Service Pack on (only Windows NT4)
			oSys.Username = js_reg_read(oReg.win,"DefaultUserName");
			oSys.Domain = js_reg_read(oReg.win,"DefaultDomainName");
			oSys.InstallDate = js_reg_read(oReg.cur,"InstallDate"); // Return a time string
		}
		return oSys;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_sendmail(iOpt,sFile,sServer,sTo,sFrom,sSubject,sMessage,bDoNotLoad,oMail,iPort){
	try{ // Needs file wsinettools.dll
		var bResult = false;
		if(arguments.length >= 7){
			if(iOpt == 1){
				if(oMail = (oMail ? oMail : js_reg_xload("wsInetTools.SMTP",sFile))){
					oMail.MailServer = sServer;
					oMail.MailPort = iPort ? iPort : 25;
					oMail.SendMail(sFrom,sTo,sSubject,sMessage);
					if(!bDoNotLoad) js_reg_regsvr("unregister",sFile);
					bResult = oMail; //
				}
				//else js.err.textarea_error.innerText += "\nFunction: js_sys_sendmail()\nErr: Could not load ActiveX Libary 'wsInetTools.SMTP'";
			}
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return bResult;
	}
}

function js_sys_pstools(sOpt,sOpt2,sPSTool,sComputer,sUser,sPass){
	try{
		var sCmd = false, bResult;
		var sCredentials = js_str_isdefined(sUser,sPass) ? " -u " + sUser + " -p " + sPass : "";
		if(!oFso.FileExists(sPSTool) || !js_sys_ping(sComputer));
		else if(sOpt == "psshutdown"){
			if(sOpt2 == "restart"){
				sCmd = "\"" + sPSTool + "\" -f -r -c -m \"System will be restarted in 20 seconds. Abort if needed.\"" + sCredentials + " \\\\" + sComputer
			}
			else if(sOpt2 == "logoff"){
				sCmd = "\"" + sPSTool + "\" -f -o " + sCredentials + " \\\\" + sComputer
			}
			var iResult = oWsh.Run(sCmd,oReg.hide,true)
			bResult = iResult == 3000 ? true : false
		}
		else if(sOpt == "psexec"){
			if(sOpt2 == "show"){
				sCmd = "%comspec% /k \"" + sPSTool + "\" \\\\" + sComputer + sCredentials + " cmd"
				bResult = oWsh.Run(sCmd,oReg.show,false)
			}
			if(sOpt2 == "showexit"){
				sCmd = "%comspec% /k \"" + sPSTool + "\" \\\\" + sComputer + sCredentials + " cmd & exit"
				bResult = oWsh.Run(sCmd,oReg.show,false)
			}
			else {
				sCmd = "%comspec% /c \"" + sPSTool + "\" \\\\" + sComputer + sCredentials + " cmd"
				bResult = oWsh.Run(sCmd,oReg.hide,true)
			}
		}
		return sCmd && bResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_getipaddress(sComputer){
	try{
		var sFile = oReg.temp + "\\" + oFso.GetTempName()
		if(js_sys_isipaddress(sComputer)) return sComputer
		else {
			var sCmd = "%comspec% /c nbtstat -R >nul & ping " + sComputer + " -n 1 -l 32 | find /i \"TTL\" 2>nul 1>" + sFile;
			if(oWsh.Run(sCmd,oReg.hide,true) == 0 && oFso.FileExists(sFile)){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault), sIPaddress;
				if(js_str_isdefined(sIPaddress = oFile.ReadAll())){ // Reply from 10.35.5.12: bytes=32 time=30ms TTL=125
					oFile.Close()
					if(sIPaddress.search(js.ore.ipaddress) > -1){
						return RegExp.$1 + "." + RegExp.$2 + "." + RegExp.$3 + "." + RegExp.$4;
					}
				}
			}
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile);
	}
}

function js_sys_getipnslookup(sHostname,sDNSServer){
	try{
		var sFile = oReg.temp + "\\j" + oFso.GetTempName()
		if(js_sys_isipaddress(sHostname)) return sHostname
		else {
			var sCmd = "%comspec% /c ipconfig /flushdns >nul & nslookup -type=A " + sHostname + " " + sDNSServer + " | find /v /i \"" + sDNSServer + "\" | find /i \"address\">" + sFile;
			if(oWsh.Run(sCmd,oReg.hide,true) == 0 && oFso.FileExists(sFile)){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault), sIPaddress;
				if(js_str_isdefined(sIPaddress = oFile.ReadAll())){ //
					oFile.Close()
					if(sIPaddress.search(js.ore.ipaddress) > -1){
						return RegExp.$1 + "." + RegExp.$2 + "." + RegExp.$3 + "." + RegExp.$4;
					}
				}
			}
		}
		return js_sys_getipaddress(sHostname);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile);
	}
}

function js_sys_getipnetmask(sComputer,sIpAddress){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName
		if((sIpAddress && js_sys_isipaddress(sIpAddress)) || (sIpAddress = js_sys_getipaddress(sComputer))){
			var aIP = sIpAddress.split(".");
			return aIP[0] + "." + aIP[1] + "." + aIP[2] + ".0";
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_getnetsession(sComputer,sPsExec){
	try{
		var aSession = false;
		sComputer = sComputer ? sComputer : oWno.ComputerName
		if(js_sys_ping(sComputer)){
			var sFile = oReg.temp + "\\" + oFso.GetTempName()
			if(!js_sys_localhost(sComputer)) sCmd = "%comspec% /c \"" + sPsExec + "\" \\\\" + sComputer + " net session | find /i \"Windows\" > " + sFile;
			else sCmd = "%comspec% /c net session | find /i \"Windows\" > " + sFile;
			if(oWsh.Run(sCmd,oReg.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
				var oRe = new RegExp("\\\\([0-9\.]+)[ \s]+([a-z0-9]+)[ \s]+([a-z0-9 ]+) ([0-9])[ \s]+([0-9:]+).*","i"); //	MAC Address = 00-B0-D0-D5-EA-64
				aSession = new Array()
				while(!oFile.AtEndOfStream){
					sLine =  oFile.ReadLine();
					if(oRe.exec(sLine)){
						var o = new Object()
						o.computer = RegExp.$1
						o.username = RegExp.$2
						o.clienttype = RegExp.$3
						o.opens = RegExp.$4
						o.idletime = RegExp.$5
						aSession.push(o);
					}
				}
				oFile.Close();
			}
			io_file_delete(sFile);
		}
		return aSession;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_getcomputername(sIpAddress){
	try{
		var sComputer = false;
		var sFile = oReg.temp + "\\" + oFso.GetTempName()
		if(!js_sys_isipaddress(sIpAddress) && !(sIpAddress = js_sys_getipaddress(sIpAddress))) return false
		//else if(!js_sys_ping(sIpAddress)) return false
		var sCmd = "%comspec% /c nbtstat -R >nul & nbtstat -A " + sIpAddress + " | find /i \"<03>\" 2>nul 1>" + sFile;
		var sCmd2 = "%comspec% /c nbtstat -A " + sIpAddress + " | find /i \"<00>\" 2>nul 1>" + sFile;
		if(oWsh.Run(sCmd,oReg.hide,true) == 0 || oWsh.Run(sCmd2,oReg.hide,true) == 0){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			var oRe = new RegExp("[ \t]+([a-z0-9\-]+)[ \t]+.*","i"); //	MAC Address = 00-B0-D0-D5-EA-64
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				if(oRe.exec(sLine)){
					sComputer = (RegExp.$1).toUpperCase();
					break;
				}
			}
			oFile.Close();
		}
		return sComputer;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile);
	}
}

function js_sys_getcomputername2(sIpAddress){
	try{
		var sComputer = false;
		var sFile = oReg.temp + "\\" + oFso.GetTempName()
		var sCmd = "%comspec% /c nbtstat -R >nul & ping -a -n 1 -l 32 " + sIpAddress + " | find /i \"pinging\" 2>nul 1>" + sFile;
		if(oWsh.Run(sCmd,oReg.hide,true) == 0){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			var oRe = new RegExp("pinging[ \t]+([a-z0-9\-\.]+)[ \t].+","ig"); // Pinging kawari.dnb.no [10.20.16.28] with 32 bytes of data:
			sComputer = js_str_trim(oFile.ReadAll())
			oFile.Close();
			sComputer = sComputer.search(oRe) > -1 ? (RegExp.$1).toLowerCase() : false
		}
		return sComputer;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile);
	}
}

function js_sys_getipmacaddress(sComputer,sCountry){
	try{
		var sMacAddress = false;
		if(js_sys_ping(sComputer)){
			sCountry = typeof(sCountry) == "string" ? sCountry : ""
			var sFile = oReg.temp + "\\" + oFso.GetTempName()
			var sMacAddressStr = sCountry.search(/no/ig) > -1 ? "MAC-adresse" : "MAC Address";
			var sCmd = "%comspec% /c nbtstat -a " + sComputer + " | find /i \"" + sMacAddressStr + "\" > " + sFile;
			if(oWsh.Run(sCmd,oReg.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
				var oRe = new RegExp(".*(= )([a-z0-9\-]{17}).*","i"); //	MAC Address = 00-B0-D0-D5-EA-64
				while(!oFile.AtEndOfStream){
					var sLine =  oFile.ReadLine();
					if(oRe.exec(sLine)){
						sMacAddress = (RegExp.$2).toUpperCase();
						var oRe = new RegExp("[\-]*","g");
						sMacAddress = sMacAddress.replace(oRe,"");
						break;
					}
				}
			}
			oFile.Close();
			io_file_delete(sFile);
		}
		return sMacAddress;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_getdhcpdomain(sDHCPServer,sComputer){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName;
		var sDomain = false, sNetMask
		if(sNetMask = js_sys_getipnetmask(sComputer)){
			var sFile = oReg.temp + "\\" + oFso.GetTempName()
			var sCmd = "%comspec% /c netsh dhcp server \\\\" + sDHCPServer + " scope " + sNetMask + " dump | find /i \"dhcp\" | find /i \"optionvalue 15\" >" + sFile
			oWsh.Run(sCmd,oReg.hide,true)
			// Dhcp Server 10.59.252.5 Scope 10.59.204.0 set optionvalue 15 STRING "Mosjoen.helgelandsb.lan"
			var s = js_str_trim(io_file_read(sFile,false,true))
			var oRe = new RegExp(".+ STRING \"([a-z\.]{6,})\"","ig");
			io_file_delete(sFile);
			if(oRe.exec(s)){
				sDomain = RegExp.$1
				if(sDomain.search(/([a-z]+)\.([a-z]+)\.([a-z]{2,5})/ig) > -1) sDomain = RegExp.$1; // Mosjoen
			}
		}
		return js_str_capitalize(sDomain);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_getipconfig(sServer,bPing){
	try{
		if(bPing && !js_sys_ping(sServer))	return false
		var oIpconfig = new Object(), aKeys
		oIpconfig.IpAddress = false
		var o = new class_mgnt_registry(null,sServer)
		oIpconfig.DNSSearchList = o.getValue(oReg.tcpip,"SearchList")
		oIpconfig.Domain = o.getValue(oReg.tcpip,"Domain")
		oIpconfig.Hostname = (o.getValue(oReg.tcpip,"Hostname")).toUpperCase()
		oIpconfig.IPRouter = o.getValueDWORD(oReg.tcpip,"IPEnableRouter") ? true : false
		oIpconfig.ReservedPorts = o.getValueMULTI(oReg.tcpip,"ReservedPorts")
		if(aKeys = o.getKeys(oReg.tcpip + "\\DNSRegisteredAdapters")){
			for(var j = 0, len = aKeys.length; j < len; j++){
				if(aKeys[j].substring(0,1) == "{"){ // {0711509E-13F3-412A-BD03-8962B397D622}
					if(oIpconfig.IpAddress = o.getValueMULTI(oReg.tcpip_int + "\\" + aKeys[j],"IPAddress")){
						oIpconfig.PrimaryDomainName = o.getKeys(oReg.tcpip + "\\DNSRegisteredAdapters\\" + aKeys[j])
						oIpconfig.NameServers = o.getValue(oReg.tcpip_int + "\\" + aKeys[j],"NameServer")
						oIpconfig.SubnetMask = o.getValue(oReg.tcpip_int + "\\" + aKeys[j],"SubnetMask")
						oIpconfig.DhcpServer = o.getValue(oReg.tcpip_int + "\\" + aKeys[j],"DhcpServer")
						oIpconfig.DefaultGateway = o.getValue(oReg.tcpip_int + "\\" + aKeys[j],"DefaultGateway")
						oIpconfig.EnableDHCP = o.getValueDWORD(oReg.tcpip_int + "\\" + aKeys[j],"EnableDHCP") ? true : false
						oIpconfig.LeaseObtained = o.getValueDWORD(oReg.tcpip_int + "\\" + aKeys[j],"LeaseObtainedTime")
						oIpconfig.LeaseTerminate = o.getValueDWORD(oReg.tcpip_int + "\\" + aKeys[j],"LeaseTerminatesTime")
						oIpconfig.WINSservers = o.getValueMULTI(oReg.nbt_int + "\\Tcpip_" + aKeys[j],"NameServerList")
					}
				}
			}
		}
		return (oIpconfig.IpAddress ? oIpconfig : false);
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_ipconfig(iOpt,sComputer,sDrv,sDomain,sUser,sPass){
	try{
		if(iOpt != 2 || js_sys_localhost(sComputer)){ // Local
			var sFile = oReg.temp + "\\" + oFso.GetTempName()
			var sCmd = "%comspec% /c ipconfig /all > " + sFile;
			oWsh.Run(sCmd,oReg.hide,true);
		}
		else if(iOpt == 2){ // Remote
			var sCmd = "rclient \\\\" + sComputer + " /L:" + sDomain + "\\" + sUser + " " + sPass + " /R \"ipconfig /all > c:\\temp\\trl_ipconfig.txt\"";
			oWsh.Run(sCmd,oReg.hide,true);
			var sFile = sDrv + "\\temp\\js_sys_ipconfig.txt";
		}

		var oIP = new Object();
		var sLine, sItem, sValue, iDns = 2, iWin = 1;
		var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUse);
		var oRe = new RegExp("([a-zA-Z\-\xE5\xE6\xF8\xC5\xC6\xD8]*)\.{1,12}:(.*)","i"); // isSearches also å,æ,ø,Å,Æ,Ø (Norwegian characters)

		while(!oFile.AtEndOfStream){
			sLine = ( js_str_rem((sOrgLine = oFile.ReadLine())," ") ).toString();
			if(oRe.exec(sLine)){
				sItem = (RegExp.$1).toLowerCase(), sValue = RegExp.$2;
				if(sItem == "hostname" || sItem == "vertsnavn") oIP.hostname = sValue;
				else if(sItem == "dnsservers" || sItem == "dns-servere") oIP.dnsserver1 = sValue;
				else if(sItem == "primarydnssuffix") oIP.dnssuffix = sValue;
				else if(sItem == "dnssuffixsearchlist") oIP.dnssearch = sValue;
				else if(sItem == "nodetype" || sItem == "nodetype") oIP.nodetype = sValue;
				else if(sItem == "iproutingenabled" || sItem == "ip-rutingmuliggjort") oIP.iprouting = sValue;
				else if(sItem == "winsproxyenabled" || sItem == "winsproxymuliggjort") oIP.winsproxy = sValue;
				else if(sItem == "description" || sItem == "beskrivelse") oIP.description = oRe.exec(sOrgLine) ? RegExp.$2 : sValue;
				else if(sItem == "physicaladdress" || sItem == "fysiskadresse") oIP.macaddress = sValue;
				else if(sItem == "dhcpenabled" || sItem == "dhcpmuliggjort") oIP.dhcpenabled = sValue;
				else if(sItem == "autoconfigurationenabled") oIP.autoconfig = sValue;
				else if(sItem == "ipaddress" || sItem == "ip-adresse") oIP.ipaddress = sValue;
				else if(sItem == "subnetmask" || sItem == "nettverksmaske") oIP.netmask = sValue;
				else if(sItem == "defaultgateway" || sItem == "standardgateway") oIP.gateway = sValue;
				else if(sItem == "dhcpserver" || sItem == "dhcp-server") oIP.dhcpserver = sValue;
				else if(sItem == "primarywinsserver" || sItem == "rwins-server") oIP["winsserver" + iWin++] = sValue; // primærwins-server
				else if(sItem == "secondarywinsserver" || sItem == "rwins-server") oIP["winsserver" + iWin++] = sValue; // sekundærwins-server: fungerar ikke riktigt, problem med norska tecken
				else if(sItem == "leaseobtained" || sItem == "leiemottatt") oIP.leaseobtain = oRe.exec(sOrgLine) ? RegExp.$2 : sValue;
				else if(sItem == "leaseexpires" || sItem == "leieforfaller") oIP.leaseexpire = oRe.exec(sOrgLine) ? RegExp.$2 : sValue;
				else {
					sLine = (js_str_rem(sLine,"\t")).split(/[\n\r]/g);
					if(sLine[0].substring(0,24) == "netbiosresolutionusesdns" || sLine[0].substring(0,27) == "netbios-oppløsningbrukerdns"){ // still a bug with norwegian
						oIP.netbios = sLine[0].substring(25,28); // A bug because of no dots
					}
					else if(sItem == "sningbrukerdns") oIP.netbios = sValue;
				}
			}
			else {
				sLine = js_str_rem(sLine,"\t");
				if(js_sys_isipaddress(sLine)) oIP["dnsserver" + iDns++] = sLine;
			}
		}
		oFile.Close();
		io_file_delete(sFile);
		return oIP;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_wins(){
	try{
		var sFile = oReg.temp + "\\" + oFso.GetTempName()
		var sCmd = "%comspec% /c for /f \"tokens=12\" %a in ('ipconfig /all ^| find /i \"Primary WINS\"') do echo %a>" + sFile
		var sResult = false
		if(oWsh.Run(sCmd,oReg.hide,true) == 0){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			sResult = (oFile.ReadAll()).replace(/\n|\t|\r/ig,"")
			oFile.close()
		}
		return sResult
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile)
	}
}

function js_sys_netsh(sOpt,sNetsh,sWINS,sName){
	try{
		var sResult = false
		var sFile = oReg.temp + "\\" + oFso.GetTempName()
		sWINS = typeof(sWINS) == "string" ? sWINS : js_sys_wins()
		
		oWsh.Run("%comspec% /c nbtstat -R>nul & nbtstat -RR>nul & ipconfig /flusgdns>nul",oReg.hide,true)
		if(sOpt == "getipaddress"){ // Access rights is need on the WINS server
			
			/*
			***You have Read and Write access to the server 10.68.68.68***

			Name                  : AB37360        [03h]
			NodeType              : 3
			State                 : ACTIVE
			Expiration Date       : 18. oktober 2006 08:03:06
			Type of Rec           : UNIQUE
			Version No            : 0 3b7ad5f
			RecordType            : DYNAMIC
			IP Address            : 10.195.4.83
			Command completed successfully.
			*/
			var sCmd1 = "%comspec% /c for /f \"tokens=4\" %a in ('netsh wins server \\\\" + sWINS + " show name " + sName
						+ " 03 ^| find /i \"IP Address\"') do echo %a>" + sFile
			var sCmd2 = "%comspec% /c for /f \"tokens=4\" %a in ('\"" + sNetsh + "\"  wins server \\\\" + sWINS + " show name " + sName
						+ " 03 ^| find /i \"IP Address\"') do echo %a>" + sFile
			if(oWsh.Run(sCmd1,oReg.hide,true) == 0 || oWsh.Run(sCmd2,oReg.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);				
				sResult = (oFile.ReadAll()).replace(/\n|\t|\r/ig,"") // 10.195.4.83
				oFile.close()
			}
		}
		else if(sOpt == "getcomputer"){			
			var sCmd1 = "%comspec% /c for /f \"tokens=4\" %a in ('netsh wins server \\\\" + sWINS + " show name " + sName
						+ " 03 ^| find /i \"IP Address\"') do for /f \"tokens=2\" %i in ('nslookup %a ^| find /i \"Name:\"') do echo %i>" + sFile
			var sCmd2 = "%comspec% /c for /f \"tokens=4\" %a in ('\"" + sNetsh + "\"  wins server \\\\" + sWINS + " show name " + sName
						+ " 03 ^| find /i \"IP Address\"') do for /f \"tokens=2\" %i in ('nslookup %a ^| find /i \"Name:\"') do echo %i>" + sFile
			if(oWsh.Run(sCmd1,oReg.hide,true) == 0 || oWsh.Run(sCmd2,oReg.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
				sResult = (oFile.ReadAll()).replace(/\n|\t|\r/ig,"") // sm78273.dnbnor.net
				sResult = (sResult.replace(/([a-z0-9\-]+).*$/ig,"$1")).toUpperCase()
				oFile.close()
			}
		}
		return sResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		io_file_delete(sFile)
	}
}

function js_sys_ping(sComputer,bOpt){
	try{
		oWsh.Run("%comspec% /c nbtstat -R & nbtstat -RR & ipconfig /flushdns",oReg.hide,false)
		if(bOpt == 1){ // Works in all circumstances but shows the window
			var oRe = new RegExp("TTL=","ig");
			var oPingExec = oWsh.Exec("ping -n 1 -l 10 -w 1 -f -i 1 " + sComputer);
			var sPingStream = oPingExec.StdOut.ReadAll();
			return (oRe.test(sPingStream));
		}
		else { // Works only when "unknown host". Won't work if computer is unavailable (offline,turned off) but defined in DNS since ping returns "Request time out" with errorcode 0.
			var sCmd = "%comspec% /c ping " + sComputer + " -n 1 -l 100 -w 1 -f -i 1 | find /i \"TTL\" >nul";
			return (oWsh.Run(sCmd,oReg.hide,true) == 0);
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_admintools(sOpt,sServer){
	try{
		sServer = sServer ? sServer : oWno.ComputerName
		if(sOpt == "msinfo32"){
			oWsh.Run("winmsd /computer " + sServer,oReg.hide,false);
		}
		else if(sOpt == "eventvwr"){
			oWsh.Run("%systemroot%\\system32\\eventvwr.exe \\\\" + sServer,oReg.hide,false)
		}
		else if(sOpt == "compmgmt"){
			oWsh.Run("%systemroot%\\system32\\compmgmt.msc -s /computer:\\\\" + sServer,oReg.hide,false)
		}
		else if(sOpt == "explorer"){
			oWsh.Run("%comspec% /c start \\\\" + sServer + "\\c$",oReg.hide,false)
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_getbytes(iBytes){
	try{
		if(!js_str_isnumber(iBytes)) return 0
		else if(iBytes >= 1000000000) return Math.round(iBytes/1047000000) + " GB";
		else if(1000000000 > iBytes && iBytes >= 1000000) return parseInt(iBytes/1047000) + " MB";
		else if(1000000 > iBytes && iBytes >= 1000) return parseInt(iBytes/1024) + " KB";
		else return parseInt(iBytes) + " Bytes";
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_netaddgroup(sServer,sPsExec,sUser,sPass,sLocalGroup,sGlobalGroup){
	try{
		sLocalGroup = sLocalGroup ? sLocalGroup : "Administrators"
		if(!js_str_isdefined(sServer,sPsExec,sUser,sPass)) return false
		var sCmd = "\"" + sPsExec + "\" \\\\" + sServer + " -u " + sUser + " -p " + sPass + " net localgroup " + sLocalGroup + " " + sGlobalGroup + " /ADD"
		if((r=oWsh.Run(sCmd,oReg.hide,true)) != 0 && r != 2) return false; // r=0 update OK. r=2: Already updated
		return true
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_netuse(sDrive,sShare,sDomain,sUser,sPassword,bUpdateProfile){
	try{
		if(oFso.DriveExists(sDrive)){
			if(oDrive = io_drive_info(sDrive)){
				var sCmd = "%comspec% /c subst " + sDrive + " /D";
				if(oWsh.Run(sCmd,oReg.hide,true) != 0 && oDrive.type != "network") return false;
			}
			else return false;
			try{ oWno.RemoveNetworkDrive(sDrive,true);}
			catch(ee){}
		}

		// Removes IPC connection is exists
		var oRe1 = new RegExp("\\\\\\\\([0-9a-zA-Z_]{9,12})\\\\([0-9a-zA-Z_]+)","ig");
		var oRe2 = new RegExp("\\\\\\\\([a-zA-Z]{4}[0-9]{3}N)\\\\([0-9a-zA-Z_]+)","ig");
		if(oRe1.exec(sShare) || oRe2.exec(sShare)){
			var sComputer = RegExp.$1;
			var sCmd = "%comspec% /c net use \\\\" + sComputer + "\\IPC$ /del /y";
			oWsh.Run(sCmd,oReg.hide,true);
		}
		else return false;

		bUpdateProfile = bUpdateProfile ? true : false;
		try{
			if(js_sys_ping(sComputer)){
				oWno.MapNetworkDrive(sDrive,sShare,bUpdateProfile,sDomain + "\\" + sUser,sPassword);
				return true;
			}
			return false;
		}
		catch(ee){ // This is mainly for Troll R: shares, which are not recognized as real shares
			js.err.description = ee.description;
		}

	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_netuseipc(iOpt,sComputer,sDomain,sUser,sPassword,iPing){
	try{
		if(js_sys_localhost(sComputer)) return true;
		else if(iOpt == 1 || iOpt == "connect"){
			if(js_sys_ping(sComputer,iPing)){
				js_sys_netuseipc(2,sComputer);
				sUser = sDomain ? sDomain + "\\" + sUser : sUser;
				var sCmd1 = "%comspec% /c net use \\\\" + sComputer + "\\IPC$ /u:" + sUser + " " + sPassword;
				var sCmd2 = "%comspec% /c net use \\\\" + sComputer + "\\IPC$";
				if((oWsh.Run(sCmd1,oReg.hide,true) != 0) && (oWsh.Run(sCmd2,oReg.hide,true) != 0)) return false;
			}
			else return false;
		}
		else if(iOpt == "adminshare"){
			if(js_sys_ping(sComputer,iPing)){
				sUser = sDomain ? sDomain + "\\" + sUser : sUser;
				var sCmd1 = "%comspec% /c net use \\\\" + sComputer + "\\admin$";
				var sCmd2 = "%comspec% /c net use \\\\" + sComputer + "\\admin$ /u:" + sUser + " " + sPassword;
				return ((oWsh.Run(sCmd1,oReg.hide,true) == 0) || (oWsh.Run(sCmd2,oReg.hide,true) == 0));
			}
			return false;
		}
		else if(iOpt == 2){ // Delete map
			// Removes any existing credentials, if exist
			oWsh.Run("%comspec% /c net use \\\\" + sComputer + "\\admin$ /del /y",oReg.hide,true);
			oWsh.Run("%comspec% /c net use \\\\" + sComputer + "\\IPC$ /del /y",oReg.hide,false);
		}
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_netadmin(sComputer,sUser,sPass){
	try{
		if(js_sys_localhost(sComputer)) return true;
		else if(js_sys_ping(sComputer)){
			if(oWsh.Run("%comspec% /c dir /ad /b \\\\" + sComputer + "\\c$ >nul",oReg.hide,true) != 0){
				var up = (typeof(sUser) == "string" && typeof(sPass) == "string") ? " /u:" + sUser + " " + sPass : ""
				if(oWsh.Run("%comspec% /c mode con cols=80 lines=10 & color 87 & Title Admin connection on: " + sComputer + " & net use \\\\" + sComputer + "\\admin$ " + up + " /p:no",oReg.show,true) != 0){
					return false
				}
			}
			return true
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_version(){
	try{
		var sFile = oReg.temp + "\\" + oFso.GetTempName()
		var sCmd = "%comspec% /c ver > " + sFile;
		oWsh.Run(sCmd,oReg.hide,true);
		var oSys = false;
		var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
		oFile.SkipLine();
		var sLine = oFile.ReadLine();
		if(sVer = sLine.isSearch(/Windows NT versjon 4.0/i)){
			oSys = new Object()
			oSys.osmodel = "NT";
			oSys.osname = "Windows NT";
			oSys.country = "NO";
			oSys.language = "Norwegian (Bokmål)";
			oSys.version = "4.0";
		}
		else if(sVer = sLine.isSearch(/Windows NT version 4.0/i)){
			oSys = new Object()
			oSys.osmodel = "NT";
			oSys.osname = "Windows NT";
			oSys.country = "EN";
			oSys.language = "English";
			oSys.version = "4.0";
		}
		else if(sVer = sLine.isSearch(/Microsoft Windows 2000 \[Version 5.00.2195\]/i)){ // Server
			oSys = new Object()
			oSys.osmodel = "2K";
			oSys.osname = "Microsoft Windows 2000 Profesional|Server";
			oSys.country = "EN";
			oSys.language = "English"
			oSys.version = "5.0";
		}
		else if(sVer = sLine.isSearch(/Microsoft Windows XP \[Version 5.1.2600\]/i)){ // Profesional
			oSys = new Object()
			oSys.osmodel = "XP";
			oSys.osname = "Microsoft Windows XP Professional";
			oSys.country = "EN";
			oSys.language = "English"
			oSys.version = "5.1";
		}
		if(oSys){
			oSys.servpack = js_reg_read(oReg.cur,"CSDVersion");
			oSys.procident = js_reg_envget("SYSTEM","PROCESSOR_IDENTIFIER");
			sPath = (js_reg_envget("SYSTEM","Path")).toLowerCase();
			oSys.path = js_str_replace(sPath,";","; ");
		}
		oFile.Close();
		io_file_delete(sFile);
		return oSys;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_service(sOpt,sOpt2,sComputer,sService){
	try{
		var bResult = false;
		sComputer = sComputer ? sComputer : oWno.ComputerName
		if(sOpt == "rpc"){
			if(sOpt2 == "restart"){
				// The "Computer Browser" service is depended on the "server" service
				// Stopping "server" will also stop Computer Browser
				// Stopping "workstation" will disconect all mappings
				var sCmd = "%comspec% /c net stop netlogon /y >nul & net stop server /y >nul";
				if(oWsh.Run(sCmd,oReg.hide,true) != 0) bResult = false;
				sCmd = "%comspec% /c /y net start browser /y >nul & net start netlogon /y >nul";
				if(oWsh.Run(sCmd,oReg.hide,true) != 0) bResult = false;
			}
		}
		else {
			if(js_str_isdefined(sOpt2,sService)){
				sAction = sOpt2 == "stop" ? "stop" : "start"
				sCmd = "%comspec% /c sc \\\\" + sComputer + " " +  sAction + " \"" + sService + "\" >nul"
				bResult = (oWsh.Run(sCmd,oReg.hide,true) == 1);
			}
		}
		return bResult
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		return bResult;
	}
}

function js_sys_dir(iDir,iPversion){
	try{
		var sDir = false;
		/*
		http://msdn.microsoft.com/library/default.asp?url=/library/en-us/graphics/hh/graphics/prtinst_976v.asp
		66000 = C:\WINNT\System32\spool\DRIVERS\W32X86\2
		66001 = C:\WINNT\System32\spool\PRTPROCS\W32X86\2
		66002 = system32 (always)
		66003 = C:\WINNT\System32\spool\PRTPROCS\W32X86\2

		http://msdn.microsoft.com/library/default.asp?url=/library/en-us/infguide/hh/infguide/infguide_5ew5.asp
		00 Null LDID. This LDID can be used to create a new LDID
		01 Source Drive:\pathname
		10 Machine directory (Maps to the Windows directory on a Server-Based Setup.)
		11 System directory
		12 IOSubsys directory
		13 Command directory
		17 INF Directory
		18 Help directory
		20 Fonts
		21 Viewers
		22 VMM32
		23 Color directory
		24 Root of drive containing the Windows directory
		25 Windows directory
		26 Guarenteed boot device for Windows (Winboot)
		28 Host Winboot
		30 Root directory of the boot drive, Ex: DefaultDestDirs=30,bin ; Direct copies go to Boot:\bin
		31 Root directory for Host drive of a virtual boot drive
		*/
		switch(iDir){
			case '66000' : case '52' : {
				sDir = (iPversion == 2) ? oReg.w32 : oReg.w32_2
				break;
			}
			case '66001' : {
				sDir = oReg.w32;
				break;
			}
			case '11' : case '12' : case '66002' : {
				sDir = oReg.system32;
				break;
			}
			case '17' : { sDir = oReg.windir + "\\inf"; break;}
			case '18' : { sDir = oReg.windir + "\\help"; break;}
			case '25' : { sDir = oReg.windir; break;}
			default : break;
		}
		return sDir;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}


function js_sys_rundll(iOpt,sType,sCommand){ // http://www.dx21.com/SCRIPTING/RUNDLL32/REFGUIDE.ASP
	try{
		if(iOpt == 1){
			if(sType == "PrintersFolder") var sCmd = "%comspec% /c RunDLL32.EXE SHELL32.DLL,SHHelpShortcuts_RunDLL PrintersFolder";
		}
		else {
			var sCmd = "%comspec% /c " + sCommand;
		}
		oWsh.Run(sCmd,oReg.show,true);
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_sys_rclient(iOpt,sComputer,sDrv,aFiles,sTmpFolder,sCommand,bMap,sDomain,sUser,sPass,bReboot){
	try{ // Needs rclient and shutdown
		if(iOpt == 1 && js_sys_ping(sComputer)){
			if(!js_sys_localhost(sComputer)){
				var sShare = "\\\\" + sComputer + "\\C$";
				var sFolder = sDrv + "\\temp\\" + sTmpFolder;
				sCommand = "cd /d " + "C:\\temp\\" + sTmpFolder + " & " + sCommand;

				// Mapping drive if requested
				if(bMap && !js_sys_netuse(sDrv,sShare,sDomain,sUser,sPass)) return false;
				if(!oFso.DriveExists(sDrv)) js_sys_netuse(sDrv,sShare,sDomain,sUser,sPass);

				// Managing folder
				js_io_delfolderrec(sFolder);
				oFso.CreateFolder(sFolder);

				// Copying files to the temporily folder
				for(var i = 0, len = aFiles.length; i < len; i++){
					if(oFso.FileExists(aFiles[i])){
						oFso.CopyFile(aFiles[i],sFolder + "\\",true);
					}
				}

				// Running rclient
				sCommand = "rclient \\\\" + sComputer + " /l:" + sDomain + "\\" + sUser + " " + sPass + " /R \"" + sCommand + "\"";
				oWsh.Run(sCommand,oReg.hide,true);

				// Deleting files & drive
				js_io_delfolderrec(sFolder);
				js_io_drivesdelete(2,sDrv);
			}
			else {
				sPath = oFso.GetParentFolderName(aFiles[0]);
				oWsh.Run("%comspec% /c cd /d " + sPath + " & " + sCommand,oReg.hide,true);
			}

			// Rebooting
			if(bReboot){
				// do something
			}
		}
		else if(iOpt == 2){
			sCmd = "%comspec% /c " + sCommand;
		}
		return true;
	}
	catch(e){
		js_io_delfolderrec(sFolder);
		js_log_error(2,e);
		return false;
	}
}

function js_io_msdos(sCommand,iMaxwait){
	try { // Ex sCommand: %comspec% /c dir
		iMaxwait = js_str_isnumber(iMaxwait) ? iMaxwait : 30000; // 30 seconds
		var iSleep = js.time.mil3
		js.err.number = js.err.description = null
		var oCommand = new Object();
		oCommand.ExitCode = 9999
		var oExec = oWsh.Exec(sCommand);
		oCommand.ProcessID = oExec.ProcessID; // (Read-Only) the process ID (PID) of the run process
		for(var iWait = iSleep; true; iWait += iSleep){
			if(oExec.Status == 0 || iWait > iMaxwait){
				break;
			}
			js_tme_sleep(iSleep);
		}
		oCommand.ExitCode = oExec.ExitCode; // (Read-Only) exit code returned by the run process
		oCommand.message = js_str_trim(oExec.StdOut.ReadAll())
		oExec.Terminate()
		oCommand.message_err = js_str_trim(oExec.StdErr.ReadAll())
		return oCommand;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		oExec = null;
	}
}


function js_reg_xload(xName,xFile,bRegistered){
	try{
		return new ActiveXObject(xName);
	}
	catch(e){
		if(!bRegistered && oFso.FileExists(xFile)){
			if(js_reg_regsvr("register",xFile,false)){
				return js_reg_xload(xName,xFile,true);
			}
		}
		else js_log_print("log_result","Unable to locate file: " + xFile);
		return false;
	}
}

function js_reg_regsvr(sOpt,sFile){
	try{
		if(oFso.FileExists(sFile = oFso.GetAbsolutePathName(sFile))){
			var sMode = sOpt == "register" ? " " : " /u ";
			return (oWsh.Run("%comspec% /c regsvr32 /s" + sMode + "\"" + sFile + "\"",oReg.hide,true) == 0);
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////
/////// WSH SCRIPT CREATE FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////


function js_wsh_create(sFile,sTitle,sStart,sFunctionStream){
	try{
		var s = "///////////\n//// " + sTitle + "\n\n";
		s = s + 'var oFso = new ActiveXObject("Scripting.FileSystemObject");\n'
		s = s + 'var oWsh = new ActiveXObject("WScript.Shell");\n'
		s = s + 'var oWno = new ActiveXObject("WScript.Network");\n'
		s = s + 'var oReg = new js_reg_object();\n\n';
		s = s + 'if(' + sStart + ') WScript.Quit(0);\nelse WScript.Quit(-1);';
		io_file_append(sFile,s,true);

		var sFunction, oFile = oFso.OpenTextFile(sFile,oReg.append,true,oReg.TristateUseDefault);
		oFile.WriteBlankLines(1);
		oFile.WriteLine(js_reg_object);
		for(var i = 3, len = arguments.length; i < len; i++){
			if(sFunction = arguments[i]){
				oFile.WriteBlankLines(2);
				oFile.WriteLine(sFunction);
			}
		}
		oFile.Close();
		return true;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_wsh_sleep(){
	var Args = WScript.arguments;
	if(Args.length == 1){
		WScript.Sleep(Args(0));
	}
	WScript.Quit(0);
}


////////////////////////////////////////////////////////////////////////////////////////////////
/////// LOG/ERROR/DEBUG FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////

function js_log_print(sOpt,sStatus,sBreak){
	try{
		if(typeof sStatus != "string") return;
		if(sOpt == "log_result"){
			for(var iWait = 0, iSleep; js.log_res_isactive; ){ // my version of synchronization (in Java language)
				if(js_log_print_maxwait < iWait) break;
				iSleep = js_str_randomize(5,300), iWait += iSleep;
				js_tme_sleep(iSleep);
			}
			js.log_res_isactive = true;
			var d = (new Date()).formatDateTime(), aStatus
			if((aStatus = sStatus.split(/[\n\r]/g)) && (len = aStatus.length) > 1){
				for(var i = 0, s = ""; i < len; i++){
					s = s + "\n[" + d + "] " + aStatus[i];
				}
				sStatus = s
			}
			else sStatus = "[" + d + "] " + sStatus
			sBreak = sBreak ? sBreak : "\n"
			
			var oTag = false
			if(js.log.logfile) io_file_append(js.err.logfile,sStatus,false,true);
			if(js.err.textarea_name) {
				oTag = document.all[js.err.textarea_name]
				oTag.innerText += sBreak + sStatus
			}
			else if(js.err.textarea_result){
				oTag = js.err.textarea_result
				oTag.innerText += sBreak + sStatus
			}
			else WScript.Echo(sStatus);

			js.log_res_isactive = false
			if(oTag && js.log.scroll) js.log.scroll(oTag)
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function js_log_command(bOpt,msg,noreturn,logfile){
	try{
		logfile = logfile ? logfile : oReg.temp + "\\" + oFso.GetTempName()
		var title = "Script Program Status", copycon = "copy con " + logfile;
		var titlecon = title + " - " + copycon;
		if(bOpt){
			if(!oWsh.AppActivate(title)){
				io_file_delete(logfile);
				var cmd = "%comspec% /k title " + title;
				oWsh.Run(cmd,oReg.show,false);
				js_app_activate(title);
				oWsh.SendKeys("~@echo off~");
				oWsh.SendKeys("color 87~");
				oWsh.SendKeys("mode con cols=75 lines=40~cls~");
				oWsh.SendKeys(copycon + "~");
				//oWsh.SendKeys(js_tme_date() + ", " + js_tme_time() + "~");
				oWsh.SendKeys("~");
				oWsh.SendKeys(msg + "~");
			}
			else {
				oWsh.SendKeys(msg);
				if(!noreturn) oWsh.SendKeys("~");
			}
		}
		else {
			if(oWsh.AppActivate(titlecon)){
				oWsh.SendKeys("^z~exit~");
				if(oFso.FileExists(logfile)){
					//oWsh.Run("notepad " + logfile,oReg.show,true);
					io_file_delete(logfile);
				}
			}
		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}


/*
JScript syntax errors are errors that result when the structure of one of your JScript
statements violates one or more of the grammatical rules of the JScript scripting language.
JScript syntax errors occur during the program compilation stage, before the program has begun to be executed.

Error Number Description
1019 Can't have 'break' outside of loop
1020 Can't have 'continue' outside of loop
1030 Conditional compilation is turned off
1027 'default' can only appear once in a 'switch' statement
1005 Expected '('
1006 Expected ')'
1012 Expected '/'
1003 Expected ':'
1004 Expected ';'
1032 Expected '@'
1029 Expected '@end'
1007 Expected ']'
1008 Expected '{'
1009 Expected '}'
1011 Expected '='
1033 Expected 'catch'
1031 Expected constant
1023 Expected hexadecimal digit
1010 Expected identifier
1028 Expected identifier, string or number
1024 Expected 'while'
1014 Invalid character
1026 Label not found
1025 Label redefined
1018 'return' statement outside of function
1002 Syntax error
1035 Throw must be followed by an expression on the same source line
1016 Unterminated comment
1015 Unterminated string constant

*/

function js_log_error(iOpt,oErr,sFunc){
	try{
		if(iOpt == 1){
			sFunc = typeof(sFunc) == "string" ? sFunc + ":\n" : ""
			try{
				alert(sFunc + ":\n" + oErr.description);
			}
			catch(ee){
				WScript.Echo(sFunc + oErr.getErrorText());
			}
		}
		else if(iOpt == 2){
			if(js.err.textarea_error) js.err.textarea_error.innerText += "\n\n" + oErr.getErrorText(false,sFunc);
			else if(js.err.divtable_error) js.err.divtable_error("log_error",oErr.getErrorObject(false,sFunc));
			else js_log_error(1,oErr,sFunc);
			if(js.err.logfile) io_file_append(js.err.logfile,oErr.getErrorText(false,sFunc),false,true); // Append to error log file
			
			if(js.log.scroll && js.err.textarea_error) js.log.scroll(js.err.textarea_error)
		}
	}
	catch(e){
		js_log_error(1,e);
	}
}

