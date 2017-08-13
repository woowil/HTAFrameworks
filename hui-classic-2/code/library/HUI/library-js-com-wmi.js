// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-com-wmi.js
Purpose:  Development script for WMI Repository
Author:   Woody Wilson
Created:  2003-08


Description:
JScript library scripting functions used for WSH files or HTA Applications.
 
Usage:
Need library-js.js

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
	try{
		//var oLastErr = new ActiveXObject("WbemScripting.SWbemLastError");
		var oLoc = new ActiveXObject("WbemScripting.SWbemLocator");
	}
	catch(e2){alert("WMI is not supported on this system. Can't load support-js-wmi.js")}
}
catch(e1){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
	try{
		//var oLastErr = new ActiveXObject("WbemScripting.SWbemLastError");
		var oLoc = new ActiveXObject("WbemScripting.SWbemLocator");
	}
	catch(e2){
		sMsg = "WMI is not supported on this system. Can't load support-";
		try{ alert(sMsg);}
		catch(e3){WScript.Echo(sMsg);}
	}
}

var LIB_NAME    = "Library WMI";
var LIB_VERSION = "2.2";
var LIB_FILE    = oFso.GetAbsolutePathName("library-js-com-wmi.js")

var wmi = new wmi_object();

function wmi_object(){
	try{
		this.version = new js_objectversion(LIB_NAME,LIB_FILE,LIB_VERSION)
	}
	catch(ee){}
	this.iIdentity = 2;
	this.iImpersonate = 3;
	this.iDelegate = 4;
	this.FlagReturnImmediately = 0x10; // Used in ExecQuery
	this.FlagForwardOnly = 0x20; // Used in ExecQuery
	this.FlagUseAmendedQualifiers = 0x20000; // Enabling this flag tells WMI to return the entire managed resource blueprint (class definition) rather than just the local definition

	// WMI Registry - root\default:StdRegProv
	this.HKCR = 0x80000000; //HKEY_CLASSES_ROOT
	this.HKCU = 0x80000001; //HKEY_CURRENT_USER
	this.HKLM = 0x80000002; //HKEY_LOCAL_MACHINE
	this.HKU = 0x80000003; //HKEY_USERS
	this.HKCC = 0x80000005; //HKEY_CURRENT_CONFIG
	
	this.env = "SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment";
	
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
	
	this.err = new Object();
	this.service = null;
	this.service_cimv2 = null;
}

function wmi_wbem_check(){
	try{
		GetObject("winmgmts:");
		return true;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_wbem_service(sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags){
	try{ // http://www.microsoft.com/technet/scriptcenter/scrguide/sas_wmi_vzbp.asp
		//http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/swbemservices.asp
		wmi_log_clear();
		sComputer = typeof(sComputer) == "string" ? sComputer : null;
		sNameSpace = typeof(sNameSpace) == "string" ? sNameSpace : "root\\CIMV2";		
		var oService = false;
		if(sComputer == null || js_sys_localhost(sComputer)){// Local Connection
			//sImpersonation = typeof(sImpersonation) == "string" ? sImpersonation : "Impersonate";
			//sAuthentication = typeof(sAuthentication) == "string" ? sAuthentication : "Connect"; // Connect: Authenticates the credentials of the client only when the client establishes a relationship with the server.
			//oService = wmi_wbem_getobject(sComputer,sNameSpace,sImpersonation,sAuthentication);
			oService = oLoc.ConnectServer(null,sNameSpace,null,null);
		}
		else { // Remote Connection			
			sUser = typeof(sUser) == "string" ? sUser : null;
			sPass = typeof(sPass) == "string" ? sPass : null;
			sLocale = typeof(sLocale) == "string" ? sLocale : "MS_409"; //MS_409 is American English
			sImpersonation = typeof(sImpersonation) == "number" ? sImpersonation : 3;
			// http://msdn2.microsoft.com/en-us/library/aa393972.aspx
			sAuthentication = typeof(sAuthentication) == "number" ? sAuthentication : 4; // Pkt(4): Authenticates that all data received is from the expected client.
			// Either NTLM {authority=ntlmdomain:DomainName} or Kerberos :{authority=kerberos:DomainName\ServerName}
			// If sUser=DOMAIN\username then sAuthority must be null
			sAuthority = sAuthority ? sAuthority : null;
			iPrivilege = typeof(iPrivilege) == "number" ? iPrivilege : 7; // ???: All Privileges, 7: Security
			iSecurityFlags = typeof(iSecurityFlags) == "number" ? iSecurityFlags : 0; // Connection timeout: 0 is indefinitely, 128 sec is max 
			try{
				oService = oLoc.ConnectServer(sComputer,sNameSpace,sUser,sPass,sLocale,sAuthority,iSecurityFlags,null);
			}
			catch(ee){
				oService = oLoc.ConnectServer(sComputer,sNameSpace,null,null,sLocale,sAuthority,iSecurityFlags,null);
			}
			oService.Security_.ImpersonationLevel = sImpersonation;
			oService.Security_.AuthenticationLevel = sAuthentication;
			
			if(iPrivilege == -1){
				// Adds all privileges
				for(var p = 1; p < 27; p++){
					try{
						oService.Security_.Privileges.Add(p); // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wbemprivilegeenum.asp
					}
					catch(ee){}
				}
			}
			else oService.Security_.Privileges.Add(iPrivilege);
		}
		if(sNameSpace.toLowerCase() == "root\\cimv2") wmi.service_cimv2 = oService
		wmi.service = oService;
		return oService;
	}
	catch(e){
		wmi.err.error = e;
		wmi.err.code = (e.number & 0xFFFF0000);
		if(ee.description == "Not Found"){
			js_str_popup("WBEM Repository might be corrupt on system " + sComputer + "\nSolution:\nnet stop winmgnt\nren \WINNT\system32\wbem\Repository Repository_Bad\nnet start winmgnt")
		}
		switch(e.number) {
			case 462 :
			case -2146827826 : // The remote server machine does not exist or is unavailable
				wmi.err.number = 10;
				wmi.err.description = "Host does not exist or is not available";
				break;
			case 429 :
			case -2146827859 : // ActiveX component can't create object
				wmi.err.number = 20;
				wmi.err.description = "Host is missing WMI capabilities";
				break;
			case -2147217405 : // Lacking admim permissions to connect to host
				wmi.err.number = 30;
				wmi.err.description = "Missing permissions to connect to this host";
				break;
			case -2146828218 : // "Permission denied" - DCOM not enabled error message
				wmi.err.number = 40;
				wmi.err.description = "DCOM not enabled on host: " + sComputer;
				break;
			case -2147418111 : // Wierd shit error (only number, no message). Happens on 2 Win2K Prof with CheckPoint's client...
				wmi.err.number = 50;
				wmi.err.description = "Wierd one (ERROR# -2147418111)...";
				break;
			case -2147217394 : // Got this when trying to connect to namespace root\MicrosoftHIS on a system where HIS is not installed
				wmi.err.number = 60;
				wmi.err.description = "Unable to connect to namespace: " + sNameSpace;
				break;
			default : // Whatever else the cause may be...
				wmi.err.number = e.number;
				wmi.err.description = e.description;
				break;
		}
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_wbem_getobject(sComputer,sNameSpace,sImpersonation,sAuthentication){
	try{ // Local connection only
		sComputer = sComputer ? sComputer : ".";
		sNameSpace = sNameSpace ? sNameSpace : "root\\CIMV2";
		sImpersonation = sImpersonation ? sImpersonation : "Impersonate";
		sAuthentication = sAuthentication ? sAuthentication : "Connect";
		//alert("winmgmts:{impersonationLevel=" + sImpersonation + ",authenticationLevel=" + sAuthentication + ",(Security)}!\\\\" + sComputer + "\\" + sNameSpace)
		var oService = GetObject("winmgmts:{impersonationLevel=" + sImpersonation + ",authenticationLevel=" + sAuthentication + ",(Security)}!\\\\" + sComputer + "\\" + sNameSpace);
		wmi.service = oService;
		return oService;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_wbem_getrefresher(oService,sClass,sServer,sNameSpace,sUser,sPass){
	try{
		oService = oService ? oService : wmi_wbem_service(sServer,sNameSpace,sUser,sPass)
		var oRefresher = new ActiveXObject("WbemScripting.SWbemRefresher")
		var oColItems = oRefresher.AddEnum(oService,sClass).objectSet
		oRefresher.Refresh()
		return oColItems;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_class_namespaces(bAll,oService,sComputer,sRoot,sUser,sPass,sImpersonate,sLocale,sAuthentication,sAuthority,iPrivilege,aNameSpaces){
	try{
		aNameSpaces = !aNameSpaces ? new Array() : aNameSpaces;
		sRoot = typeof(sRoot) == "string" ? sRoot : "root"
		oService = oService ? oService : wmi_wbem_service(sComputer,sRoot,sUser,sPass,sImpersonate,sLocale,sAuthentication,sAuthority,iPrivilege);
		if(!oService) return aNameSpaces
		var oNameSpace = oService.InstancesOf("__NAMESPACE");
		for(var oEnum = new Enumerator(oNameSpace); !oEnum.atEnd(); oEnum.moveNext()){
			var oItem = oEnum.item();
			aNameSpaces.push(sRoot + "\\" + oItem.Name);
			if(bAll) aNameSpaces.concat(wmi_class_namespaces(bAll,null,sComputer,sRoot + "\\" + oItem.Name,sUser,sPass,sImpersonate,sLocale,sAuthentication,sAuthority,iPrivilege,aNameSpaces))
		}
		return aNameSpaces.sort(); // do not sort here... take time
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

/*
********************************************************************
* 
*
* Fetch all the classes in the currently selected namespace, and
* populate the keys of a dictionary object with the names of all
* dynamic (non-association) classes.
*/

function wmi_class_enumall(oService,bWin32,bSystem,bCim,sServer,sNameSpace,sUser,sPass){
	try{
		var aClassPath, aClass = new Array();
		oService = oService ? oService : wmi_wbem_service(sServer,sNameSpace,sUser,sPass)
		var oSubClass = oService.SubclassesOf();
		var dQual, c
		bSystem = bSystem ? true : false;
		bCim = bCim ? true : false;
		bWin32 = bWin32 ? true : false;
		for(var oItem = "", oEnum = new Enumerator(oSubClass); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			dQual = wmi_obj_qualifiername(oItem);
			if(dQual.Exists("dynamic") && (sClassPath = oItem.Path_.Class)){ // The 
				aClassPath = sClassPath.split("_");
				if(aClassPath[0] == "" || aClassPath[0] == null) aClassPath[0] = "__"; 
				if(!bSystem && aClassPath[0] == "__");
				else if(!bCim && aClassPath[0].toUpperCase() == "CIM");
				else if(!bWin32 && aClassPath[0].toUpperCase() == "WIN32");
				else {
					if(!aClass[aClassPath[0]]){
						aClass[aClassPath[0]] = new Array();	
						//aClass[aClassPath[0]].push(aClassPath[0]);
						//js_tme_sleep(js.time.mil3);
					}
					aClass[aClassPath[0]].push(sClassPath);
				}				
			}
			dQual.RemoveAll();
		}
		// Sorts the array
		for(c in aClass){
			aClass[c].sort();
		}
		return aClass;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_class_enumgenerate(oService,sClass,sLanguage,sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags){
	try{
		oService = oService ? oService : wmi_wbem_service(sComputer,sNameSpace,sUser,sPass,sImpersonate,sLocale,sAuthentication,sAuthority,iPrivilege);
		var oClass = oService.Get(sClass);
		var sHtml = ""
		sComputer = sComputer ? sComputer : oWno.ComputerName
		sUser = sUser ? sUser : null;
		sPass = sPass ? sPass : null;
		sNameSpace = sNameSpace ? sNameSpace : "root\\CIMV2"
		sLocale = sLocale ? sLocale : "MS_409"; //MS_409 is American English
		sImpersonation = sImpersonation ? sImpersonation : 3;
		sAuthentication = sAuthentication ? sAuthentication : 4; // Pkt(4): Authenticates that all data received is from the expected client.
		// Either NTLM {authority=ntlmdomain:DomainName} or Kerberos :{authority=kerberos:DomainName\ServerName}
		// If sUser=DOMAIN\username then sAuthority must be null
		sAuthority = sAuthority ? sAuthority : null; 
		iPrivilege = iPrivilege ? iPrivilege : 7; // ???: All Privileges, 7: Security
		iSecurityFlags = iSecurityFlags ? iSecurityFlags : 0; // Connection timeout: 0 is indefinitely, 128 sec is max 
		bLocal = js_sys_localhost(sComputer)
		if(sLanguage.toLowerCase() == "jscript"){ // JScript			
			sHtml = sHtml + "// Made by: Woody Wilson\n// Generated by: " + document.title;
			sHtml = sHtml + "\n// " + js.disclaimer;
			sHtml = sHtml + "\n// http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_classes.asp";			
			var h = "\n\nWScript.Echo(\"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\");\n\n";
			sHtml = sHtml + h.replace(/\\/g,"\\\\");
			if(bLocal){
				sHtml = sHtml + "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\");";
				sHtml = sHtml + "\n\nfunction wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sImpersonation,sAuthentication){";
			}
			else{
				sHtml = sHtml + "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\");";
				sHtml = sHtml + "\n\nfunction wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags){";
			}
			
			sHtml = sHtml + "\n\ttry{\n\t\tvar sHtml, sHtmlAll = \"\";";
			sHtml = sHtml + "\n\t\tsComputer = sComputer ? sComputer : \".\";";
			sHtml = sHtml + "\n\t\tsNameSpace = sNameSpace ? sNameSpace : \"root\\\\CIMV2\";";
			
			if(bLocal){
				sHtml = sHtml + "\n\t\tsImpersonation = sImpersonation ? sImpersonation : \"Impersonate\";";
				sHtml = sHtml + "\n\t\tsAuthentication = sAuthentication ? sAuthentication : \"Connect\";";
				sHtml = sHtml + "\n\t\tvar oService = GetObject(\"winmgmts:{impersonationLevel=\" + sImpersonation + \",authenticationLevel=\" + sAuthentication + \",(Security)}!\\\\\\\\\" + sComputer + \"\\\\\" + sNameSpace);";
			}
			else {
				sHtml = sHtml + "\n\t\tsUser = sUser ? sUser : null;";
				sHtml = sHtml + "\n\t\tsPass = sPass ? sPass : null;";				
				sHtml = sHtml + "\n\t\tsLocale = sLocale ? sLocale : \"MS_409\";";
				sHtml = sHtml + "\n\t\tsImpersonation = sImpersonation ? sImpersonation : 3;";
				sHtml = sHtml + "\n\t\tsAuthentication = sAuthentication ? sAuthentication : 4;";
				sHtml = sHtml + "\n\t\tsAuthority = sAuthority ? sAuthority : null;"
				sHtml = sHtml + "\n\t\tiPrivilege = iPrivilege ? iPrivilege : 7;"
				sHtml = sHtml + "\n\t\tiSecurityFlags = iSecurityFlags ? iSecurityFlags : 128;"
				
				sHtml = sHtml + "\n\n\t\tvar oLocator = new ActiveXObject(\"WbemScripting.SWbemLocator\");";
				sHtml = sHtml + "\n\t\tvar oService = oLocator.ConnectServer(sComputer,sNameSpace,sUser,sPass,sLocale,sAuthority,iSecurityFlags,null);";
				sHtml = sHtml + "\n\t\toService.Security_.ImpersonationLevel = sImpersonation;";
				sHtml = sHtml + "\n\t\toService.Security_.AuthenticationLevel = sAuthentication;";
				
				sHtml = sHtml + "\n\t\tif(iPrivilege == -1){"
				sHtml = sHtml + "\n\t\t\tfor(var p = 1; p < 27; p++){"
				sHtml = sHtml + "\n\t\t\t\ttry{"
				sHtml = sHtml + "\n\t\t\t\t\toService.Security_.Privileges.Add(p); // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wbemprivilegeenum.asp"
				sHtml = sHtml + "\n\t\t\t\t}"
				sHtml = sHtml + "\n\t\t\t\tcatch(ee){}"
				sHtml = sHtml + "\n\t\t\t}\n\t\t}"
				sHtml = sHtml + "\n\t\telse oService.Security_.Privileges.Add(7);";
			}
			sHtml = sHtml + "\n\t\tvar oColItems = oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48);";
			sHtml = sHtml + "\n\n\t\tfor(var oItem, i = 1, d, v, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){";
			sHtml = sHtml + "\n\t\t\toItem = oEnum.item(), sHtml = \"\\n [INSTANCE-\" + i + \"]\";";
			
			for(var oItem, oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
					/*
					sHtml = sHtml + "\n\t\t\tif(oItem." + oItem.Name + " != null){ // VBArray object";
						sHtml = sHtml + "\n\t\t\t\tvar a" + oItem.Name + " = new VBArray(oItem." + oItem.Name + ").toArray();";
						sHtml = sHtml + "\n\t\t\t\tfor(var j = 0; j < a" + oItem.Name + ".length; j++){";
							sHtml = sHtml + "\n\t\t\t\t\tsHtml += \"\\n  " + oItem.Name + "-\" + (j+1) + \":  \" + a" + oItem.Name + "[j];";
						sHtml = sHtml + "\n\t\t\t\t}";
						sHtml = sHtml + "\n\t\t\t}";
					sHtml = sHtml + "\n\t\t\telse sHtml += \"\\n  " + oItem.Name + ": undefined\";";
					*/					
					sHtml = sHtml + "\n\t\t\tsHtml = sHtml + ((v=wmi_class_vbarray(oItem." + oItem.Name + ",'" + oItem.Name + "')) ? v.stream : \"\\n  " + oItem.Name + ": \" + \"undefined\");";
				}
				else if(oItem.CIMType == 101){ // If it is a DateTime object
					/*
					sHtml = sHtml + "\n\t\t\tif(oItem." + oItem.Name + " != null){";
						sHtml = sHtml + "\n\t\t\t\tvar d = oItem." + oItem.Name + ";";
						sHtml = sHtml + "\n\t\t\t\tvar sDateTime = d.substring(0,4) + \"-\" + d.substring(4,6) + \"-\" + d.substring(6,8) + \" \" + d.substring(8,10) + \":\" + d.substring(10,12) + \":\" + d.substring(12,14);";
						sHtml = sHtml + "\n\t\t\t\tsHtml += \"\\n  " + oItem.Name + ": \" + sDateTime;";
					sHtml = sHtml + "\n\t\t\t}";
					sHtml = sHtml + "\n\t\t\telse sHtml += \"\\n  " + oItem.Name + ": undefined\";";
					*/
					sHtml = sHtml + "\n\t\t\tsHtml += \"\\n  " + oItem.Name + ": \" + ((d=wmi_class_date(oItem." + oItem.Name + ")) ? d : \"undefined\");";
				}
				else sHtml = sHtml + "\n\t\t\tsHtml += \"\\n  " + oItem.Name + ": \" + oItem." + oItem.Name + ";";
			}
			
			sHtml = sHtml + "\n\t\t\tsHtmlAll += sHtml;"
			sHtml = sHtml + "\n\t\t\tWScript.Echo(sHtml);";
			sHtml = sHtml + "\n\t\t\tWScript.Sleep(" + js.time.mil1 + ");";
			sHtml = sHtml + "\n\t\t}\n\t}";
			sHtml = sHtml + "\n\tcatch(e){";
			sHtml = sHtml + "\n\t\tWScript.Echo(\"Error:\\nNumber: \" + (e.number & 0xFFFF) + \"\\nDescription: \" + e.description + \" - Unable to connect to system. WMI may not be installed.\");";
			sHtml = sHtml + "\n\t\treturn false;";
			sHtml = sHtml + "\n\t}";
			sHtml = sHtml + "\n\tfinally{";
			sHtml = sHtml + "\n\t\treturn sHtmlAll;";
			sHtml = sHtml + "\n\t}";
			sHtml = sHtml + "\n}\n\n";
			
			sHtml = sHtml + wmi_class_date + "\n\n" + wmi_class_vbarray;
			
			sHtml = sHtml + "\n\nWScript.Echo(\"\\nD O N E!\")\n";
		}
		else if(sLanguage.toLowerCase() == "vbscript"){ //VBscript
			sHtml = sHtml + "' Made by: Woody Wilson\n' Generated by: " + document.title;
			sHtml = sHtml + "\n' " + js.disclaimer;
			sHtml = sHtml + "\n' http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_classes.asp"
			sHtml = sHtml + "\n\nWScript.Echo \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\"\n\n";	
			if(bLocal){
				sHtml = sHtml + "wmi_" + sClass.toLowerCase() + " \"" + sComputer + "\",\"" + sNameSpace + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\"";
				sHtml = sHtml + "\n\nSub wmi_" + sClass.toLowerCase() + "(strComputer,strNameSpace,strImpersonation,strAuthentication)";
			}
			else{
				sHtml = sHtml + "wmi_" + sClass.toLowerCase() + " \"" + sComputer + "\",\"" + sNameSpace + "\""
				sHtml = sHtml + "\n\nSub wmi_" + sClass.toLowerCase() + "(strComputer,strNameSpace,strUser,strPass,strImpersonation,strLocale,strAuthentication,strAuthority,intPrivilege,intSecurityFlags)";
			}
			
			sHtml = sHtml + "\n\tOn Error Resume Next\n\tErr.Clear";
			sHtml = sHtml + "\n\n\tstrHtml = \"\" : i = 0";
			sHtml = sHtml + "\n\tIf IsEmpty(strComputer) Or IsNull(strComputer) Then strComputer = \".\" End If";
			sHtml = sHtml + "\n\tIf IsEmpty(strNameSpace) Or IsNull(strNameSpace) Then strNameSpace = \"root\\CIMV2\" End If";
			
			if(bLocal){
				sHtml = sHtml + "\n\tIf IsEmpty(strImpersonation) Or IsNull(strImpersonation) Then strImpersonation = \"Impersonate\" End If";
				sHtml = sHtml + "\n\tIf IsEmpty(strAuthentication) Or IsNull(strAuthentication) Then strAuthentication = \"Connect\" End If";
				sHtml = sHtml + "\n\tSet objService = GetObject(\"winmgmts:{impersonationLevel=\" & strImpersonation & \",authenticationLevel=\" & strAuthentication & \",(Security)}!\\\\\" & strComputer & \"\\\" &  strNameSpace)";
			}
			else {
				sHtml = sHtml + "\n\tIf IsEmpty(strUser) Or IsNull(strUser) Then strUser = null End If";
				sHtml = sHtml + "\n\tIf IsEmpty(strPass) Or IsNull(strPass) Then strPass = null End If";
				sHtml = sHtml + "\n\tIf IsEmpty(strLocale) Or IsNull(strLocale) Then strLocale = \"MS_409\" End If";				
				sHtml = sHtml + "\n\tIf Not IsNumeric(intImpersonation) Then intImpersonation = 3 End If";
				sHtml = sHtml + "\n\tIf Not IsNumeric(strAuthentication) Then strAuthentication = 4 End If";
				sHtml = sHtml + "\n\tIf IsEmpty(strAuthority) Or IsNull(strAuthority) Then strAuthority = Nothing End If";
				sHtml = sHtml + "\n\tIf Not IsNumeric(intSecurityFlags) Then intSecurityFlags = 128 End If";
				sHtml = sHtml + "\n\tIf Not IsNumeric(intPrivilege) Then intPrivilege = 7 End If";
				
				sHtml = sHtml + "\n\n\tSet objLocator = CreateObject(\"WbemScripting.SWbemLocator\")";
				sHtml = sHtml + "\n\tSet objService = objLocator.ConnectServer(strComputer,strNameSpace,strUser,strPass,strLocale,strAuthority,intSecurityFlags)";
				sHtml = sHtml + "\n\tobjService.Security_.ImpersonationLevel = intImpersonation";
				sHtml = sHtml + "\n\tobjService.Security_.AuthenticationLevel = strAuthentication";
				sHtml = sHtml + "\n\tIf intPrivilege = -1 Then";
				sHtml = sHtml + "\n\t\tFor p = 1 To 26 Step 1";
				sHtml = sHtml + "\n\t\t\tobjService.Security_.Privileges.Add(p)";
				sHtml = sHtml + "\n\t\t\tIf Err.Number <> 0 Then";
				sHtml = sHtml + "\n\t\t\t\tErr.Clear 'Invalid Parameter";
				sHtml = sHtml + "\n\t\t\tEnd If";
				sHtml = sHtml + "\n\t\tNext";
				sHtml = sHtml + "\n\tElse";
				sHtml = sHtml + "\n\t\tobjService.Security_.Privileges.Add(intPrivilege)";
				sHtml = sHtml + "\n\tEnd If";
			}
			sHtml = sHtml + "\n\tSet objColItems = objService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)";
			sHtml = sHtml + "\n\n\tFor Each objItem in objColItems";
			sHtml = sHtml + "\n\t\ti = i + 1 : strHtml = \" [INSTANCE-\" & i & \"]\" & Chr(10)";
   			
   			for(var oItem = "", oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
						/*
						sHtml = sHtml + "\n\t\tIf Not IsNull(objItem." + oItem.Name + ") Then ' VBArray object";
							sHtml = sHtml + "\n\t\t\tFor a = LBound(objItem." + oItem.Name + ") to UBound(objItem." + oItem.Name + ")";
								sHtml = sHtml + "\n\t\t\t\tstrHtml = strHtml & \"  " + oItem.Name + "-\" & a+1 & \": \" & objItem." + oItem.Name + "(a) & Chr(10)";
							sHtml = sHtml + "\n\t\t\tNext";
						sHtml = sHtml + "\n\t\tElse strHtml = strHtml & \"  " + oItem.Name + ": undefined\" & Chr(10)";
						sHtml = sHtml + "\n\t\tEnd If";
						*/
						sHtml = sHtml + "\n\t\tstrHtml = strHtml & wmi_class_vbarray(objItem." + oItem.Name + ",\"" + oItem.Name + "\")";
				}
				else if(oItem.CIMType == 101){ // If it is a DateTime object
					/*
					sHtml = sHtml + "\n\t\tIf Not IsNull(objItem." + oItem.Name + ") Then ' DateTime object";
						sHtml = sHtml + "\n\t\t\tstrDate = Replace(objItem." + oItem.Name + ",\"*\",\"0\")";
						sHtml = sHtml + "\n\t\t\tstrDateTime = CDate(Mid(strDate,5,2) & \"/\" & Mid(strDate,7,2) & \"/\" & Left(strDate,4) & \" \" & Mid(strDate,9,2) & \":\" & Mid(strDate,11,2) & \":\" & Mid(strDate,13,2))";
						sHtml = sHtml + "\n\t\t\tstrHtml = strHtml & \"  " + oItem.Name + ": \" & strDateTime & Chr(10)";
					sHtml = sHtml + "\n\t\tElse strHtml = strHtml & \"  " + oItem.Name + ": undefined\" & Chr(10)";
					sHtml = sHtml + "\n\t\tEnd If";
					*/
					sHtml = sHtml + "\n\t\tstrHtml = strHtml & \"  " + oItem.Name + ": \" & wmi_class_date(objItem." + oItem.Name + ") & Chr(10)";
				}
				else sHtml = sHtml + "\n\t\tstrHtml = strHtml & \"  " + oItem.Name + ": \" & objItem." + oItem.Name + " & Chr(10)";
			}
			sHtml = sHtml + "\n\t\tWScript.Echo strHtml";
			sHtml = sHtml + "\n\t\tWScript.Sleep " + js.time.mil1;
			sHtml = sHtml + "\n\tNext";
			sHtml = sHtml + "\n\n\tIf Err.Number <> 0 Then";
			sHtml = sHtml + "\n\t\tWScript.Echo Err.Description, \"0x\" & Hex(Err.Number)";
			sHtml = sHtml + "\n\tEnd If\nEnd Sub";
			
			// Install Date
			sHtml = sHtml + "\n\nFunction wmi_class_date(strDate)";
			sHtml = sHtml + "\n\tstrHtml = \"\"";
			sHtml = sHtml + "\n\tIf Not IsNull(strDate) Then ' DateTime object";
			sHtml = sHtml + "\n\t\tstrDate = Replace(strDate,\"*\",\"0\")";
			sHtml = sHtml + "\n\t\tstrDateTime = Left(strDate,4) &  \"-\" & Mid(strDate,5,2) & \"-\" & Mid(strDate,7,2)  & \" \" & Mid(strDate,9,2) & \":\" & Mid(strDate,11,2) & \":\" & Mid(strDate,13,2) & \".\" & Mid(strDate,16,3)";
			sHtml = sHtml + "\n\t\tstrHtml = strHtml & strDateTime";
			sHtml = sHtml + "\n\tEnd If";
			sHtml = sHtml + "\n\twmi_class_date = strHtml";
			sHtml = sHtml + "\nEnd Function";
			
			// VBArray Function
			sHtml = sHtml + "\n\nFunction wmi_class_vbarray(objArray,strArray)";
			sHtml = sHtml + "\n\tstrHtml = \"\"";
			sHtml = sHtml + "\n\tIf IsArray(objArray) Then ' VBArray object";
			sHtml = sHtml + "\n\t\tFor a = LBound(objArray) to UBound(objArray)";
			sHtml = sHtml + "\n\t\t\tstrHtml = strHtml & \"  \" & strArray & \"-\" & a+1 & \": \" & objArray(a) & Chr(10)";
			sHtml = sHtml + "\n\t\tNext";
			sHtml = sHtml + "\n\tElse";
			sHtml = sHtml + "\n\t\tstrHtml = strHtml & \"  \" & strArray & \": \" & Chr(10)";
			sHtml = sHtml + "\n\tEnd If";
			sHtml = sHtml + "\n\twmi_class_vbarray = strHtml";
			sHtml = sHtml + "\nEnd Function";

			//sHtml += wmi_vbs_date + "\n\n" + wmi_vbs_vbarray + "\n\n";
			
			sHtml = sHtml + "\n\nWScript.Echo \"D O N E!\"\n ";
		}
		else if(sLanguage.toLowerCase() == "perlscript"){ // PerlScript (Requires ActivePerl for Windows)		
			sHtml = sHtml + "# Made by: Woody Wilson\n# Generated by: " + document.title;
			sHtml = sHtml + "\n# " + js.disclaimer;
			sHtml = sHtml + "\n# This Script requires Perl for Windows or ActivePerl at http://www.activeperl.com/Products/Language_Distributions/";
			h = "\n\nprint \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\";";
			sHtml = sHtml + h.replace(/\\/g,"\\\\");
			sHtml = sHtml + "\n\nuse Win32::OLE qw(in);";
			sHtml = sHtml + "\nuse Win32::OLE::Enum;";
			
			if(bLocal){
				sHtml = sHtml + "\n\nsub wmi_" + sClass.toLowerCase() + "($$$$){";
				sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\");";
			}
			else{
				sHtml = sHtml + "\n\nsub wmi_" + sClass.toLowerCase() + "($$$$$$$$$$){";
				sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ");";
			}			
				
			sHtml = sHtml + "\n\tmy $oBool = {0 => \"0\",1 => \"1\", \'\' => \"undefined\"};";
			sHtml = sHtml + "\n\tmy $i = 1, $sHtml = \"\";";			
			
			if(bLocal){
				sHtml = sHtml + "\n\tmy ($sComputer,$sNameSpace,$sImpersonation,$sAuthentication) = @_;";
				sHtml = sHtml + "\n\t$sComputer = SetStr($sComputer,\".\");";
				sHtml = sHtml + "\n\t$sNameSpace = SetStr($sNameSpace,\"root\\CIMV2\");".replace(/\\/g,"\\\\");
				sHtml = sHtml + "\n\t$sImpersonation = SetStr($sImpersonation,\"Impersonate\");";
				sHtml = sHtml + "\n\t$sAuthentication = SetStr($sAuthentication,\"Connect\");";
				sHtml = sHtml + "\n\tmy $oService = Win32::OLE->GetObject(\"winmgmts:{impersonationLevel=\" . $sImpersonation . \",authenticationLevel=\" . $sAuthentication . \",(Security)}!\\\\\\\\\" . $sComputer . \"\\\\\" . $sNameSpace);";
			}
			else {
				sHtml = sHtml + "\n\tmy ($sComputer,$sNameSpace,$sUser,$sPass,$sImpersonation,$sLocale,$sAuthentication,$sAuthority,$iPrivilege,$iSecurityFlags) = @_;";
				sHtml = sHtml + "\n\t$sComputer = SetStr($sComputer,\".\");";
				sHtml = sHtml + "\n\t$sNameSpace = SetStr($sNameSpace,\"root\\CIMV2\");".replace(/\\/g,"\\\\");
				sHtml = sHtml + "\n\t$sUser = SetStr($sUser,undef);";
				sHtml = sHtml + "\n\t$sPass = SetStr($sPass,undef);";
				sHtml = sHtml + "\n\t$sLocale = SetStr($sLocale,\"MS_409\");";
				sHtml = sHtml + "\n\t$sImpersonation = SetStr($sImpersonation,3);";
				sHtml = sHtml + "\n\t$sAuthority = SetStr($sAuthority,undef);";
				sHtml = sHtml + "\n\t$sAuthentication = SetStr($sAuthentication,4);";
				sHtml = sHtml + "\n\t$iSecurityFlags = SetStr($iSecurityFlags,128);";
				sHtml = sHtml + "\n\t$iPrivilege = SetStr($iPrivilege,7);";
				
				sHtml = sHtml + "\n\n\tmy $oLocator = Win32::OLE->new(\"WbemScripting.SWbemLocator\");";
				       + " || warn \"Cannot access WMI on local machine: \", Win32::OLE->LastError;";
				sHtml = sHtml + "\n\tmy $oService = $oLocator->ConnectServer($sComputer,$sNameSpace,$sUser,$sPass,$sLocale,$sAuthority,$iSecurityFlags,undef)"
				       + " || warn (\"Cannot access WMI on remote machine ($sComputer): \", Win32::OLE->LastError);";
				sHtml = sHtml + "\n\t# $oService->Security_->ImpersonationLevel = $sImpersonation; # Error: Can't modify non-lvalue subroutine call";
				sHtml = sHtml + "\n\t# $oService->Security_->AuthenticationLevel = $sAuthentication; # Error: Can't modify non-lvalue subroutine call";
				sHtml = sHtml + "\n\tif($iPrivilege eq -1){";
				sHtml = sHtml + "\n\t\tfor($i = 1; $i <= 26; $i++){";
				sHtml = sHtml + "\n\t\t\t$oService->Security_->Privileges->Add($i);";
				sHtml = sHtml + "\n\t\t}";
				sHtml = sHtml + "\n\t}";
				sHtml = sHtml + "\n\telse{";
				sHtml = sHtml + "\n\t\t$oService->Security_->Privileges->Add($iPrivilege);";
				sHtml = sHtml + "\n\t}";
			}
			sHtml = sHtml + "\n\tGetErr();";
			sHtml = sHtml + "\n\tmy $oColItems = $oService->ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48);";
			sHtml = sHtml + "\n\n\tforeach $oItem (in $oColItems){";
			sHtml = sHtml + "\n\t\t$sHtml = \"\\n\\n [INSTANCE-\" . $i++ . \"]\";";
			sHtml = sHtml + "\n\t\tif(defined($oItem)){";
			
			for(var oItem, oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
					/*
					sHtml = sHtml + "\n\t\t\t$s = 1;";
					sHtml = sHtml + "\n\t\t\tforeach $a (in @{$oItem->" + oItem.Name + "}){ # Array Object";
						sHtml = sHtml + "\n\t\t\t\t$sHtml .= \"\\n  " + oItem.Name + "-\" . $s++ . \": \" . $a;";
					sHtml = sHtml + "\n\t\t\t}";
					*/
					sHtml = sHtml + "\n\t\t\t$sHtml .= wmi_class_vbarray($oItem->" + oItem.Name + "," + oItem.Name + ");"
				}
				else if(oItem.CIMType == 11){ // If it is a boolean
					sHtml = sHtml + "\n\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . $oBool->{$oItem->" + oItem.Name + "};";
				}
				else if(oItem.CIMType == 101){ // If it is a DateTime object
					/*
					sHtml = sHtml + "\n\t\t\tif($oItem->" + oItem.Name + " ne ''){ # DateTime Object";
						sHtml = sHtml + "\n\t\t\t\t$d = $oItem->" + oItem.Name + ";";
						sHtml = sHtml + "\n\t\t\t\t$sDateTime = substr($d,0,4) . \"-\" . substr($d,4,2) . \"-\" . substr($d,6,2) . \" \" . substr($d,8,2) . \":\" . substr($d,10,2) . \":\" . substr($d,12,2) . \".\" . substr($d,15,3);";
						sHtml = sHtml + "\n\t\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . $sDateTime;";
					sHtml = sHtml + "\n\t\t\t} else {";
					sHtml = sHtml + "\n\t\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": undefined\";\n\t\t}";
					*/
					sHtml = sHtml + "\n\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . wmi_class_date($oItem->" + oItem.Name + ");";
				}
				else sHtml = sHtml + "\n\t\t\t$sHtml .= \"\\n  " + oItem.Name + ": \" . $oItem->" + oItem.Name + ";";
				
			}
			sHtml = sHtml + "\n\t\t}\n\t\tprint $sHtml;\n\t}";
			sHtml = sHtml + "\n}";
			
			// VBDate
			sHtml = sHtml + "\n\nsub wmi_class_date($){";
			sHtml = sHtml + "\n\tmy $sDateTime = \"\";";
			sHtml = sHtml + "\n\tmy($sDate) = @_;";
			sHtml = sHtml + "\n\tif(defined($sDate)){ # DateTime Object";
			sHtml = sHtml + "\n\t\t$sDate =~ s/\\*/0/g;";
			sHtml = sHtml + "\n\t\t$sDateTime = substr($sDate,0,4) . \"-\" . substr($sDate,4,2) . \"-\" . substr($sDate,6,2) . \" \" . substr($sDate,8,2) . \":\" . substr($sDate,10,2) . \":\" . substr($sDate,12,2) . \".\" . substr($sDate,15,3);";
			sHtml = sHtml + "\n\t}";
			sHtml = sHtml + "\n\treturn $sDateTime;";
			sHtml = sHtml + "\n\n}";

			// VBArray
			sHtml = sHtml + "\n\nsub wmi_class_vbarray($$){";
			sHtml = sHtml + "\n\tmy $sHtml = \"\";";
			sHtml = sHtml + "\n\tmy ($oVBarray,$sVBname) = @_;";
			sHtml = sHtml + "\n\tif(defined($oVBarray)){";
			sHtml = sHtml + "\n\t\t$s = 1;";
			sHtml = sHtml + "\n\t\tforeach $a (in @{$oVBarray}){ # Array Object";
			sHtml = sHtml + "\n\t\t\t$sHtml .= \"\\n  \" .$sVBname . \"-\" . $s++ . \": \" . $a;";
			sHtml = sHtml + "\n\t\t}";
			sHtml = sHtml + "\n\t}";
			sHtml = sHtml + "\n\telse{";
			sHtml = sHtml + "\n\t\t$sHtml .= \"\\n  \" . $sVBname . \": \";";
			sHtml = sHtml + "\n\t}";
			sHtml = sHtml + "\n\treturn $sHtml;";
			sHtml = sHtml + "\n}";
			
			// GetErr
			sHtml = sHtml + "\n\nsub GetErr{";
			sHtml = sHtml + "\n\tunless(Win32::OLE->LastError() == 0){";
			sHtml = sHtml + "\n\t\tdie(\"Error: \" . Win32::OLE->LastError());";
			sHtml = sHtml + "\n\t}";
			sHtml = sHtml + "\n}";
			
			// SetStr
			sHtml = sHtml + "\n\nsub SetStr{"
			sHtml = sHtml + "\n\tmy $sString = shift;"
			sHtml = sHtml + "\n\tmy $sDefault = shift;"
			sHtml = sHtml + "\n\tif (not defined($sString)){\n\t\t$sString = $sDefault;\n\t}"
			sHtml = sHtml + "\n\treturn $sString;"
			sHtml = sHtml + "\n}"
			
			sHtml = sHtml + "\n\n" + sHtmlInit + "\nprint \"\\n\\nD O N E!\";\n";
		}
		else if(sLanguage.toLowerCase() == "wsh"){ // WSH/WSF
			sHtml = sHtml + "<?XML version=\"1.0\" standalone=\"yes\" ?>";
			sHtml = sHtml + "\n<package>\n\t<job id=\"ILoveJS\">\n\t<?job debug=\"true\"?>";
			sHtml = sHtml + "\n\t\t<script language=\"JScript\">\n\t\t<![CDATA[\n";
      		
      		var sJSHtml = wmi_class_enumgenerate(oService,sClass,"jscript",sComputer,sNameSpace,sUser,sPass,sImpersonation,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags);
      		sHtml = sHtml + sJSHtml
      		
			sHtml = sHtml + "\n\t\t]]>\n\t\t</script>";
			sHtml = sHtml + "\n\t</job>\n</package>";
		}
		else if(sLanguage.toLowerCase() == "asp_js"){ // ASP
			sHtml = sHtml + "<%@ LANGUAGE=\"JSCRIPT\"%>";
			sHtml = sHtml + "\n\n<html>\n<!--";
			sHtml = sHtml + "\n Made by: Woody Wilson\n Generated by: " + document.title;
			sHtml = sHtml + "\n-->\n<%\n/*\n This sample illustrates the use of the WMI Scripting API within an ASP, using JScript.\n It displays information in a table for each disk on the local host. To run this sample:";
			sHtml = sHtml + "\n  1. Place it in a directory accessible to your web server\n  2. If running on NT4, ensure that the registry value:\n  HKEY_LOCAL_MACHINE\\Software\\Microsoft\\WBEM\\Scripting\\Enable for ASP is set to 1";
			sHtml = sHtml + "\n\n For Windows 2000, ensure that Anonymous Access is disabled and Windows\n Integrated Authentication is enabled for this file before running this ASP\n (this can be done by configuring the file properties using the IIS configuration snap-in).";
			sHtml = sHtml + "\n*/\n%>\n\n<head>\n<title></title>\n<style>BODY,TD{font-family:Arial;font-size:11px;padding:2px;vertical-align:top;}</style>\n</head>\n<body scroll=\"auto\">";			
			sHtml = sHtml + "\n<br><br>\n<table width=\"500\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background-color:#EEEEEE;border:1px solid #545454;\">";
			sHtml = sHtml + "\n<caption align=\"center\" style=\"background-color:black;color:white;\"><b>WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + "</b></caption>";
			sHtml = sHtml + "\n\n<%";
			sHtml = sHtml + "\ntry{";
			sHtml = sHtml + "\n\tvar sComputer = \"" + (!bLocal ? sComputer : '.') + "\";";
			sHtml = sHtml + "\n\tvar sNameSpace = \"" + sNameSpace.replace(/\\/g,"\\\\") + "\";";
			
			if(bLocal){
				sHtml = sHtml + "\n\tvar sImpersonation = \"" + sImpersonation + "\";";
				sHtml = sHtml + "\n\tvar sAuthentication = \"" + (sAuthentication ? sAuthentication : "Connect") + "\";";
				sHtml = sHtml + "\n\tvar oService = GetObject(\"winmgmts:{impersonationLevel=\" + sImpersonation + \",authenticationLevel=\" + sAuthentication + \",(Security)}!\\\\\\\\\" + sComputer + \"\\\\\" + sNameSpace);";
			}
			else {
				sHtml = sHtml + "\n\tvar sUser = \"" + (sUser ? sUser : '') + "\";";
				sHtml = sHtml + "\n\tvar sPass = \"" + (sPass ? sPass : '') + "\";";
				sHtml = sHtml + "\n\tvar sLocale = \"" + (sLocale ? sLocale : 'MS_409') + "\";";
				sHtml = sHtml + "\n\tvar sImpersonation = " + (sImpersonation ? sImpersonation : 3) + ";";
				sHtml = sHtml + "\n\tvar sAuthority = \"" + (sAuthority ? sAuthority : "ntlmdomain:" + oWno.UserDomain) + "\";";
				sHtml = sHtml + "\n\tvar sAuthentication = \"" + (sAuthentication ? sAuthentication : 4) + "\";";
				sHtml = sHtml + "\n\tvar iPrivilege = \"" + (iPrivilege ? iPrivilege : 7) + "\";";
				sHtml = sHtml + "\n\tvar iSecurityFlags = \"" + (iSecurityFlags ? iSecurityFlags : 128) + "\";";
				sHtml = sHtml + "\n\n\tvar oLocator = new ActiveXObject(\"WbemScripting.SWbemLocator\");";
				sHtml = sHtml + "\n\tvar oService = oLocator.ConnectServer(sComputer,sNameSpace,sUser,sPass,sLocale,sAuthority,iSecurityFlags,null);";
				sHtml = sHtml + "\n\toService.Security_.ImpersonationLevel = sImpersonation;";
				sHtml = sHtml + "\n\toService.Security_.AuthenticationLevel = sAuthentication;";
				sHtml = sHtml + "\n\toService.Security_.Privileges.Add(iPrivilege);";
			}
			sHtml = sHtml + "\n\tvar oColItems = oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48);";
			sHtml = sHtml + "\n\n\tfor(var oItem, i = 1, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){";
			sHtml = sHtml + "\n\t\toItem = oEnum.item();\n%>";
			sHtml = sHtml + "\n\n<tr><td colspan=\"2\">&nbsp;</td></tr>";
			sHtml = sHtml + "\n\n<tr><td colspan=\"2\" style=\"background-color:#545454;color:white\"><b>INSTANCE-<%=i%></b></td></tr>";
			
			for(var oItem, oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
					sHtml = sHtml + "\n\n<%";
					sHtml = sHtml + "\n\t\tif(oItem." + oItem.Name + " != null){";
						sHtml = sHtml + "\n\t\t\tvar a" + oItem.Name + " = new VBArray(oItem." + oItem.Name + ").toArray();";
						sHtml = sHtml + "\n\t\t\tfor(var j = 0; j < a" + oItem.Name + ".length; j++){";
						sHtml = sHtml + "\n%>";
							sHtml = sHtml + "\n\n<tr>\n\t<td>" + oItem.Name + "-<%=(j+1)%></td>\n\t<td>&nbsp;<%=a" + oItem.Name + "[j]%></td>\n</tr>";
						sHtml = sHtml + "\n\n<%";
						sHtml = sHtml + "\n\t\t\t}";
						sHtml = sHtml + "\n\t\t}";
					sHtml = sHtml + "\n\t\telse {\n%>\n\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;undefined</td>\n</tr>";
					sHtml = sHtml + "\n\n<%\n\t\t}\n%>\n";
				}
				else if(oItem.CIMType == 101){ // If it is a DateTime object
					sHtml = sHtml + "\n\n<%";
					sHtml = sHtml + "\n\t\tif(oItem." + oItem.Name + " != null){";
						sHtml = sHtml + "\n\t\t\tvar d = oItem." + oItem.Name + ";";
						sHtml = sHtml + "\n\t\t\tvar sDateTime = d.substring(0,4) + \"-\" + d.substring(4,6) + \"-\" + d.substring(6,8) + \" \" + d.substring(8,10) + \":\" + d.substring(10,12) + \":\" + d.substring(12,14);";
						sHtml = sHtml + "\n%>\n";
						sHtml = sHtml + "\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;<%=sDateTime%></td>\n</tr>";
					sHtml = sHtml + "\n\n<%";
					sHtml = sHtml + "\n\t\t}";
					sHtml = sHtml + "\n\t\telse {\n%>\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;undefined</td>\n</tr>";
					sHtml = sHtml + "\n\n<%\n\t\t}\n%>\n";
				}
				else sHtml = sHtml + "\n<tr>\n\t<td>" + oItem.Name + "</td>\n\t<td>&nbsp;<%=oItem." + oItem.Name + "%></td>\n</tr>";
			}
			sHtml = sHtml + "\n\n<%\n\t}\n}";
			sHtml = sHtml + "\ncatch(e){";
			sHtml = sHtml + "\n%>\n\nError:<br>Number: <%=(e.number & 0xFFFF)%><br>Description: <%=e.description%>";
			sHtml = sHtml + "\n\n<%\n}\n%>";
			sHtml = sHtml + "\n\n</table>\n\n</body>\n</html>\n";
		}
		else if(sLanguage.toLowerCase() == "kixstart"){ // KiXstart
			sHtml = sHtml + "; Made by: Woody Wilson\n; Generated by: " + document.title;
			sHtml = sHtml + "\n; " + js.disclaimer;
			sHtml = sHtml + "\n; http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_classes.asp"
			sHtml = sHtml + "\n\nBREAK ON";
			sHtml = sHtml + "\n\n;Redirect Outputs: kix32.exe " + sClass + ".kix $sReDirect='redirect.txt'"
			sHtml = sHtml + "\nIf IsDeclared($sReDirect) $r = ReDirectOutput($sReDirect,0) EndIf"
			sHtml = sHtml + "\n\n? \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\" + Chr(10)";
			
			if(bLocal){
				sHtml = sHtml + "\n\nwmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\")";
				sHtml = sHtml + "\n\nFunction wmi_" + sClass.toLowerCase() + "($sComputer,$sNameSpace,$sImpersonation,$sAuthentication)";
			}
			else{
				sHtml = sHtml + "\n\nwmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ")";
				sHtml = sHtml + "\n\nFunction wmi_" + sClass.toLowerCase() + "($sComputer,$sNameSpace,$sUser,$sPass,$sImpersonation,$sLocale,$sAuthentication,$sAuthority,$iPrivilege,$iSecurityFlags)";
			}			
			
			sHtml = sHtml + "\n\tDim $oService, $oLocator"
			sHtml = sHtml + "\n\t$sHtml = \"\"\n\t$i = 0";
			
			sHtml = sHtml + "\n\n\t$sComputer = SetStr($sComputer,'.')";
			sHtml = sHtml + "\n\t$sNameSpace = SetStr($sNameSpace,'root\\CIMV2')";
			if(bLocal){
				sHtml = sHtml + "\n\t$sImpersonation = SetStr($sImpersonation,'Impersonate')";
				sHtml = sHtml + "\n\t$sAuthentication = SetStr($sAuthentication,'Connect')";
				sHtml = sHtml + "\n\t$oService = GetObject(\"winmgmts:{impersonationLevel=\" + $sImpersonation + \",authenticationLevel=\" + $sAuthentication + \",(Security)}!\\\\\" + $sComputer + \"\\\" +  $sNameSpace)";
			}
			else {
				sHtml = sHtml + "\n\t$sUser = SetStr($sUser,'')";
				sHtml = sHtml + "\n\t$sPass = SetStr($sPass,'')";
				sHtml = sHtml + "\n\t$sLocale = SetStr($sLocale,'MS_409')";				
				sHtml = sHtml + "\n\t$iImpersonation = SetInt($iImpersonation,3)";
				sHtml = sHtml + "\n\t$sAuthentication = SetInt($sAuthentication,4)";
				sHtml = sHtml + "\n\t$sAuthority = SetStr($sAuthority,'')";
				sHtml = sHtml + "\n\t$iSecurityFlags = SetInt($iSecurityFlags,128)";
				sHtml = sHtml + "\n\t$iPrivilege = SetInt($iPrivilege,7)";
				
				sHtml = sHtml + "\n\n\t$oLocator = CreateObject(\"WbemScripting.SWbemLocator\")";
				sHtml = sHtml + "\n\t$oService = $oLocator.ConnectServer($sComputer,$sNameSpace,$sUser,$sPass,$sLocale,$sAuthority,$iSecurityFlags)";
				sHtml = sHtml + "\n\t$oService.Security_.ImpersonationLevel = $iImpersonation";
				sHtml = sHtml + "\n\t$oService.Security_.AuthenticationLevel = $sAuthentication";
				sHtml = sHtml + "\n\tIf $iPrivilege == -1";
				sHtml = sHtml + "\n\t\tFor $p = 1 To 26 Step 1";
				sHtml = sHtml + "\n\t\t\t$oService.Security_.Privileges.Add($p)";
				sHtml = sHtml + "\n\t\t\tGetErr(False,True)";
				sHtml = sHtml + "\n\t\tNext";
				sHtml = sHtml + "\n\tElse";
				sHtml = sHtml + "\n\t\t$oService.Security_.Privileges.Add($iPrivilege)";
				sHtml = sHtml + "\n\tEndIf";
			}			
			sHtml = sHtml + "\n\n\t$oColItems = $oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)";
			sHtml = sHtml + "\n\tGetErr()";

			sHtml = sHtml + "\n\n\tFor Each $oItem in $oColItems";
			sHtml = sHtml + "\n\t\t$i = $i + 1\n\t\t$sHtml = \" [INSTANCE-\" + $i + \"]\" + Chr(10)";
   			
   			for(var oItem = "", oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
						sHtml = sHtml + "\n\t\t$sHtml = $sHtml + wmi_class_vbarray($oItem." + oItem.Name + ",\"" + oItem.Name + "\")";
				}
				else if(oItem.CIMType == 11){ // If it is a boolean
					sHtml = sHtml + "\n\t\t$sHtml = $sHtml + \"  " + oItem.Name + ": \" + GetBoolean($oItem." + oItem.Name + ") + Chr(10)";
				}
				else if(oItem.CIMType == 101){ // If it is a DateTime $oect
					sHtml = sHtml + "\n\t\t$sHtml = $sHtml + \"  " + oItem.Name + ": \" + wmi_class_date($oItem." + oItem.Name + ") + Chr(10)";
				}
				else sHtml = sHtml + "\n\t\t$sHtml = $sHtml + \"  " + oItem.Name + ": \" + $oItem." + oItem.Name + " + Chr(10)";
			}
			sHtml = sHtml + "\n\t\t? $sHtml";
			sHtml = sHtml + "\n\tNext";
			sHtml = sHtml + "\n\tGetErr()";
			sHtml = sHtml + "\nEndFunction";
			
			// Install Date
			sHtml = sHtml + "\n\nFunction wmi_class_date($sDate)";
			sHtml = sHtml + "\n\t$sHtml = \"\"";
			sHtml = sHtml + "\n\tIf VarType($sDate) == 12 Or VarType($sDate) == 8; String or DateTime Object";
			sHtml = sHtml + "\n\t\t;$sDate = Replace($sDate,\"*\",\"0\") ;Not implemented";
			sHtml = sHtml + "\n\t\t$sDateTime = Left($sDate,4) + \"-\" + SubStr($sDate,5,2) + \"-\" + SubStr($sDate,7,2) + \" \" + SubStr($sDate,9,2) + \":\" + SubStr($sDate,11,2) + \":\" + SubStr($sDate,13,2) + \".\" + SubStr($sDate,16,3)";
			sHtml = sHtml + "\n\t\t$sHtml = $sHtml + $sDateTime";
			sHtml = sHtml + "\n\tEndIf";
			sHtml = sHtml + "\n\t$wmi_class_date = $sHtml";
			sHtml = sHtml + "\nEndFunction";
			
			// VBArray Function
			sHtml = sHtml + "\n\nFunction wmi_class_vbarray($oArray,$sArray)";
			sHtml = sHtml + "\n\t$sHtml = \"\"";
			sHtml = sHtml + "\n\tIf VarType($oArray) >= 8192 ; VBArray Object";
			sHtml = sHtml + "\n\t\tFor $a = 0 to UBound($oArray)";
			sHtml = sHtml + "\n\t\t\t$sHtml = $sHtml + \"  \" + $sArray + \"-\" + ($a+1) + \": \" + $oArray[$a] + Chr(10)";
			sHtml = sHtml + "\n\t\tNext";
			sHtml = sHtml + "\n\tElse";
			sHtml = sHtml + "\n\t\t$sHtml = $sHtml + \"  \" + $sArray + \": \" + Chr(10)";
			sHtml = sHtml + "\n\tEndIf";
			sHtml = sHtml + "\n\t$wmi_class_vbarray = $sHtml";
			sHtml = sHtml + "\nEndFunction";
			
			// SetStr Function
			sHtml = sHtml + "\n\nFunction SetStr($sString,$sDefault)"
			sHtml = sHtml + "\n\t$sString = Trim($sString)"
			sHtml = sHtml + "\n\tIf $sString = ''"
			sHtml = sHtml + "\n\t\t$sString = $sDefault"
			sHtml = sHtml + "\n\tEndIf"
			sHtml = sHtml + "\n\t$SetStr = $sString"
			sHtml = sHtml + "\nEndFunction"
			
			// SetInt
			sHtml = sHtml + "\n\nFunction SetInt($iInteger,$iDefault)"
			sHtml = sHtml + "\n\tIf VarType($iInteger) <> 3"
			sHtml = sHtml + "\n\t\t$iInteger = $iDefault"
			sHtml = sHtml + "\n\tEndIf"
			sHtml = sHtml + "\n\t$SetInt = $iInteger"
			sHtml = sHtml + "\nEndFunction"
			
			// GetBoolean
			sHtml = sHtml + "\n\nFunction GetBoolean($iBoolean)";
			sHtml = sHtml + "\n\t$sBoolean = \"False\"";
			sHtml = sHtml + "\n\tIf VarType($iBoolean) == 3 And $iBoolean == -1";
			sHtml = sHtml + "\n\t\t$sBoolean = \"True\"";
			sHtml = sHtml + "\n\tEndIf";
			sHtml = sHtml + "\n\t$GetBoolean = $sBoolean";
			sHtml = sHtml + "\nEndFunction";
			
			// GetErr
			sHtml = sHtml + "\n\nFunction GetErr(optional $bExit, optional $bClear)"
			sHtml = sHtml + "\n\tIf @ERROR <> 0"
			sHtml = sHtml + "\n\t\t? \" Number: \" + @ERROR + Chr(10) + \" Description: \" + @SERROR  + Chr(10)"
			sHtml = sHtml + "\n\t\tIf $bExit\n\t\t\tExit If(@ERROR < 0, VAL(\"&\" + Right(DecToHex(@ERROR),4)),@ERROR)\n\t\tEndIf"
			sHtml = sHtml + "\n\tEndIf"
			sHtml = sHtml + "\n\tIf $bClear @ERROR = 0 EndIf"
			sHtml = sHtml + "\nEndFunction"			
			
			sHtml = sHtml + "\n\n? \"D O N E!\"\n";
		}
		else if(sLanguage.toLowerCase() == "rexx"){ // Object rexx http://www.mindspring.com/~dave_martin/RexxFAQ.html
			sHtml = sHtml + "/* Made by: Woody Wilson\n* Generated by: " + document.title;
			sHtml = sHtml + "\n* " + js.disclaimer;
			sHtml = sHtml + "\n* This Script requires IBM Object Rexx for Windows http://www.rexx.com\n*/";
			sHtml = sHtml + "\n\nsay '[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]'";
			sHtml = sHtml + "\n\n";
			
			if(bLocal){
				sHtml = sHtml + "wmi_" + sClass.toLowerCase() + ":\nProcedure Expose sComputer sNameSpace sImpersonation sAuthentication";
				sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\")";
			}
			else{
				sHtml = sHtml + "wmi_" + sClass.toLowerCase() + ":\nProcedure Expose sComputer sNameSpace sUser sPass sImpersonation sAuthentication sLocale sAuthentication sAuthority iPrivilege iSecurityFlags";
				sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ")";
			}			
				
			sHtml = sHtml + "\n\ti = 1\n\tsHtml = \"\"";			
			
			if(bLocal){
				sHtml = sHtml + "\n\tsComputer = \".\"";
				sHtml = sHtml + "\n\tsNameSpace = \"root\\CIMV2\""
				sHtml = sHtml + "\n\toService = .OLEObject.GetObject(\"winmgmts:\\\\\"||sComputer||\"\\\"||sNameSpace)";
			}
			else {
				sHtml = sHtml + "\n\tsComputer = \".\"";
				sHtml = sHtml + "\n\tsNameSpace = \"root\\CIMV2\""
				sHtml = sHtml + "\n\toService = .OLEObject~ConnectServer(sComputer,sNameSpace)";
				
			}
			sHtml = sHtml + "\n\toColItems = oService~ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)";
			sHtml = sHtml + "\n\n\tdo oItem over oColItems";
			sHtml = sHtml + "\n\t\tsHtml = \"\\n\\n [INSTANCE-\" i++ \"]\"";
			
			for(var oItem, oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
					sHtml = sHtml + "\n\t\tsHtml = sHtml wmi_class_vbarray(oItem~" + oItem.Name + ",\"" + oItem.Name + "\")"
				}/*
				else if(oItem.CIMType == 11){ // If it is a boolean
					sHtml = sHtml + "\n\t\t\tsHtml = sHtml || \"\\n  " + oItem.Name + ": \" + oBool.{oItem~" + oItem.Name + "}";
				}*/
				else if(oItem.CIMType == 101){ // If it is a DateTime object
					sHtml = sHtml + "\n\t\tsHtml = sHtml \"\\n  " + oItem.Name + ": \" wmi_class_date(oItem~" + oItem.Name + ")";
				}
				else sHtml = sHtml + "\n\t\tsHtml = sHtml \"\\n  " + oItem.Name + ": \" oItem~" + oItem.Name + "";
				
			}
			sHtml = sHtml + "\n\t\tsay sHtml\n\treturn";
			
			// VBDate
			sHtml = sHtml + "\n\nwmi_class_date:\nProcedure oDtm";
			sHtml = sHtml + "\n\tsDateTime = \"\"";
			sHtml = sHtml + "\n\tif(oDtm[4] == 0):";
			sHtml = sHtml + "\n\t\tsDateTime = oDtm[5] + '-'";
			sHtml = sHtml + "\n\telse:";
			sHtml = sHtml + "\n\t\tsDateTime = oDtm[4] + oDtm[5] + '-'";
			sHtml = sHtml + "\n\tif(oDtm[6] == 0):";
			sHtml = sHtml + "\n\t\tsDateTime = sDateTime + oDtm[7] + '-'";
			sHtml = sHtml + "\n\telse:";
			sHtml = sHtml + "\n\t\tsDateTime = sDateTime + oDtm[6] + oDtm[7] + '-'";
			sHtml = sHtml + "\n\t\tsDateTime = oDtm[0] + oDtm[1] + oDtm[2] + oDtm[3] + sDateTime + \" \" + oDtm[8] + oDtm[9] + \":\" + oDtm[10] + oDtm[11] + \":\" + oDtm[12] + oDtm[13]";
			sHtml = sHtml + "\n\treturn sDateTime";

			// VBArray
			sHtml = sHtml + "\n\nwmi_class_vbarray:\nProcedure oVBarray sVBname";
			sHtml = sHtml + "\n\tsItem = \"\"";
			sHtml = sHtml + "\n\ts = 1";
			sHtml = sHtml + "\n\tif .NIL <> oVBarray then";
			sHtml = sHtml + "\n\t\tdo x over oVBarray"
			sHtml = sHtml + "\n\t\t\tsItem = sItem \"\\n  \" sVBname \"-\" s++ \": \" x";
			sHtml = sHtml + "\n\tend";
			sHtml = sHtml + "\n\treturn sItem";	
			
			sHtml = sHtml + "\n\n" + sHtmlInit + "\n\nsay \"\\n\\nD O N E!\"\n";
		}
		else if(sLanguage.toLowerCase() == "python"){ // Python
			sHtml = sHtml + "# Made by: Woody Wilson\n# Generated by: " + document.title;
			sHtml = sHtml + "\n# " + js.disclaimer;
			sHtml = sHtml + "\n# This Script requires Python for Windows http://www.python.org/doc";
			sHtml = sHtml + "\n\nprint \"[ WMI Class: \\\\" + sComputer + "\\" + sNameSpace + ":" + sClass + " ]\"";
			sHtml = sHtml + "\n\nimport win32com.client";
			
			if(bLocal){
				sHtml = sHtml + "\n\ndef wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sImpersonation,sAuthentication):";
				sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sImpersonation + "\",\"" + sAuthentication + "\")";
			}
			else{
				sHtml = sHtml + "\n\ndef wmi_" + sClass.toLowerCase() + "(sComputer,sNameSpace,sUser,sPass,sImpersonation,sAuthentication,sLocale,sAuthentication,sAuthority,iPrivilege,iSecurityFlags):";
				sHtmlInit = "wmi_" + sClass.toLowerCase() + "(\"" + sComputer + "\",\"" + sNameSpace.replace(/\\/g,"\\\\") + "\",\"" + sUser + "\",\"" + sPass + "\",\"" + sImpersonation + "\",\"" + sLocale + "\",\"" + sAuthentication + "\",\"" + sAuthority + "\"," + iPrivilege + "," + iSecurityFlags + ")";
			}			
				
			sHtml = sHtml + "\n\ti = 1, sHtml = \"\"";			
			
			if(bLocal){
				sHtml = sHtml + "\n\tsComputer = \".\"";
				sHtml = sHtml + "\n\tsNameSpace = sNameSpace #SetStr(sNameSpace)"
				sHtml = sHtml + "\n\toService = win32com.client.Dispatch(\"WbemScripting.SWbemLocator\")"
				sHtml = sHtml + "\n\toService = oService.ConnectServer(sComputer,sNameSpace)";
			}
			else {
				sHtml = sHtml + "\n\tsComputer = \".\"";
				sHtml = sHtml + "\n\tsNameSpace = sNameSpace #SetStr(sNameSpace)"
				sHtml = sHtml + "\n\toService = win32com.client.Dispatch(\"WbemScripting.SWbemLocator\")"
				sHtml = sHtml + "\n\toService = oService.ConnectServer(sComputer,sNameSpace)";
				
				/*
				sHtml = sHtml + "\n\t$sComputer = SetStr($sComputer,\".\");";
				sHtml = sHtml + "\n\t$sNameSpace = SetStr($sNameSpace,\"root\\CIMV2\");".replace(/\\/g,"\\\\");
				sHtml = sHtml + "\n\t$sUser = SetStr($sUser,undef);";
				sHtml = sHtml + "\n\t$sPass = SetStr($sPass,undef);";
				sHtml = sHtml + "\n\t$sLocale = SetStr($sLocale,\"MS_409\");";
				sHtml = sHtml + "\n\t$sImpersonation = SetStr($sImpersonation,3);";
				sHtml = sHtml + "\n\t$sAuthority = SetStr($sAuthority,undef);";
				sHtml = sHtml + "\n\t$sAuthentication = SetStr($sAuthentication,4);";
				sHtml = sHtml + "\n\t$iSecurityFlags = SetStr($iSecurityFlags,128);";
				sHtml = sHtml + "\n\t$iPrivilege = SetStr($iPrivilege,7);";
				
				sHtml = sHtml + "\n\n\tmy $oLocator = Win32::OLE->new(\"WbemScripting.SWbemLocator\");";
				       + " || warn \"Cannot access WMI on local machine: \", Win32::OLE->LastError;";
				sHtml = sHtml + "\n\tmy $oService = $oLocator->ConnectServer($sComputer,$sNameSpace,$sUser,$sPass,$sLocale,$sAuthority,$iSecurityFlags,undef)"
				       + " || warn (\"Cannot access WMI on remote machine ($sComputer): \", Win32::OLE->LastError);";
				sHtml = sHtml + "\n\t# $oService->Security_->ImpersonationLevel = $sImpersonation; # Error: Can't modify non-lvalue subroutine call";
				sHtml = sHtml + "\n\t# $oService->Security_->AuthenticationLevel = $sAuthentication; # Error: Can't modify non-lvalue subroutine call";
				sHtml = sHtml + "\n\tif($iPrivilege eq -1){";
				sHtml = sHtml + "\n\t\tfor($i = 1; $i <= 26; $i++){";
				sHtml = sHtml + "\n\t\t\t$oService->Security_->Privileges->Add($i);";
				sHtml = sHtml + "\n\t\t}";
				sHtml = sHtml + "\n\t}";
				sHtml = sHtml + "\n\telse{";
				sHtml = sHtml + "\n\t\t$oService.Security_->Privileges->Add($iPrivilege);";
				sHtml = sHtml + "\n\t}";
				*/
			}
			sHtml = sHtml + "\n\toColItems = oService.ExecQuery(\"Select * from " + sClass + "\",\"WQL\",48)";
			sHtml = sHtml + "\n\n\tfor oItem in oColItems:";
			sHtml = sHtml + "\n\t\tsHtml = \"\\n\\n [INSTANCE-\" + i++ + \"]\"";
			
			for(var oItem, oEnum = new Enumerator(oClass.properties_); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.IsArray){ // Only VBArrays
					sHtml = sHtml + "\n\t\tsHtml = sHtml + wmi_class_vbarray(oItem." + oItem.Name + ",\"" + oItem.Name + "\")"
				}/*
				else if(oItem.CIMType == 11){ // If it is a boolean
					sHtml = sHtml + "\n\t\t\tsHtml = sHtml + \"\\n  " + oItem.Name + ": \" + oBool.{oItem->" + oItem.Name + "}";
				}*/
				else if(oItem.CIMType == 101){ // If it is a DateTime object
					sHtml = sHtml + "\n\t\tsHtml = sHtml + \"\\n  " + oItem.Name + ": \" + wmi_class_date(oItem." + oItem.Name + ")";
				}
				else sHtml = sHtml + "\n\t\tsHtml = sHtml + \"\\n  " + oItem.Name + ": \" + oItem." + oItem.Name + "";
				
			}
			sHtml = sHtml + "\n\t\tprint sHtml";
			
			// VBDate
			sHtml = sHtml + "\n\ndef wmi_class_date(oDtm):";
			sHtml = sHtml + "\n\tsDateTime = \"\"";			
			sHtml = sHtml + "\n\tif(oDtm[4] == 0):";
			sHtml = sHtml + "\n\t\tsDateTime = oDtm[5] + '-'";
			sHtml = sHtml + "\n\telse:";
			sHtml = sHtml + "\n\t\tsDateTime = oDtm[4] + oDtm[5] + '-'";
			sHtml = sHtml + "\n\tif(oDtm[6] == 0):";
			sHtml = sHtml + "\n\t\tsDateTime = sDateTime + oDtm[7] + '-'";
			sHtml = sHtml + "\n\telse:";
			sHtml = sHtml + "\n\t\tsDateTime = sDateTime + oDtm[6] + oDtm[7] + '-'";
			sHtml = sHtml + "\n\t\tsDateTime = oDtm[0] + oDtm[1] + oDtm[2] + oDtm[3] + sDateTime + \" \" + oDtm[8] + oDtm[9] + \":\" + oDtm[10] + oDtm[11] + \":\" + oDtm[12] + oDtm[13]";
			sHtml = sHtml + "\n\treturn sDateTime";

			// VBArray
			sHtml = sHtml + "\n\ndef wmi_class_vbarray(oVBarray,sVBname):";
			sHtml = sHtml + "\n\tsItem = \"\"";
			sHtml = sHtml + "\n\ttry:";
			sHtml = sHtml + "\n\t\ts = 1";
			sHtml = sHtml + "\n\t\tfor a in oVBarray:";
			sHtml = sHtml + "\n\t\t\tsItem = sItem + \"\\n  \" + sVBname + \"-\" + s++ + \": \" + a";
			sHtml = sHtml + "\n\texcept:";
			sHtml = sHtml + "\n\t\tsItem = sItem + \"\\n  \" + sVBname + \": \"";
			sHtml = sHtml + "\n\treturn sItem";	
			
			sHtml = sHtml + "\n\n" + sHtmlInit + "\n\nprint \"\\n\\nD O N E!\"\n";
		}
		return sHtml;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}


function wmi_class_derivation(sClass,oService,sRoot){
	try{
		oService = oService ? oService : wmi_wbem_service(false,sRoot);
		var oClass = oService.Get(sClass);
		if(aDerivedTest = wmi_obj_isvbarray(oClass.Derivation_)) var aDerived = aDerivedTest.reverse();
		else var aDerived = new Array();
		var t = '<table border="0" cellspacing="0" cellpadding="0" class="cTableA">';
		t = t + '\n<caption align="left"><i>Derivation for ' + sClass + '</i></caption>';
		t = t + '\n<tr><th width="12">&nbsp;</th><th>Name</th><th>Type</th></tr>';
		t = t + '\n<tr valign="top"><td>0</td><td>|_ ' + " " + sRoot.toUpperCase() + '</td><td>ROOT</td></tr>';
		for(var i = 0, sThread = "", len = aDerived.length; i < len; i++){
			if(i == 0) sType = "Dynasty";
			else if(i == (aDerived.length-1)) sType = "SuperClass"; 
			else sType = "Class";
			sThread = sThread + "_"
			t = t + '\n<tr valign="top"><td>' + (i+1) + '</td><td>|_' + sThread + " " + aDerived[i] + '</td><td>' + sType + '</td></tr>';
		}
		t = t + '\n<tr valign="top"><td>' + (i+1) + '</td><td>|__' + sThread + " " + sClass + '</td><td>Class</td></tr>';
		t = t + '\n</table>';
		aDerived.table = t
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return aDerived;
	}
}

function wmi_class_method(sClass,oService){
	try{
		oService = oService ? oService : wmi_wbem_service();
		var aMethod = new Array();
		var oClass = oService.Get(sClass);
		
		for(var oItem, o, aInParam, oParam, oEnum = new Enumerator(oClass.Methods_); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			o = new Object()
			// Name
			o.name = oItem.Name;
			// In Parameter
			o.inparam = ""
			if(aInParam = wmi_meth_inparameter(oItem)){
				for(var j = 0, len = aInParam.length; j < len; j++){
					o.inparam = o.inparam + aInParam[j].name + " As " + aInParam[j].typestr + " <" + aInParam[j].cimtype + ">\n";
					aInParam[j] = null
				}
			}
			o.inparam = o.inparam == "" ? "" : (o.inparam).replace(/(.*)\n$/g,"$1")
			// Out Parameter
			oParam = wmi_meth_outparameter(oItem);
			if(oParam.name) o.outparam = oParam.name + " As " + oParam.typestr + " <" + oParam.cimtype + ">"
			else o.outparam = ""
			// Description
			o.description = aInParam ? aInParam.desc : ""
			// Qualifies
			o.qualifiers = wmi_obj_qualifier(oItem)
			aMethod.push(o)
			aInParam = oParam = null
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return aMethod;
	}
}

function wmi_class_method_old(sClass,oService){
	try{
		oService = oService ? oService : wmi_wbem_service();
		var aMethod = new Array();
		var oClass = oService.Get(sClass);
		var oMethod = oClass.Methods_;
		
		aMethod.table = '<table width="100%" border="0" cellspacing="1" cellpadding="0" class="cTable1">';
		aMethod.table += '\n<caption align="left"><i>Methods for ' + sClass + '</i></caption>';
		aMethod.table += '\n<tr align="left"><th width="10">&nbsp;</th><th>Name</th><th>Description</th></tr>';
		
		for(var oItem, i = 0, oEnum = new Enumerator(oMethod); !oEnum.atEnd(); i++, oEnum.moveNext()){
			oItem = oEnum.item();
			aMethod[i] = new Object()
			aMethod[i].name = oItem.Name;
			
			// OutParameter
			oOutParam = wmi_meth_outparameter(oItem);
			aMethod[i].type = oOutParam.type1;
			aMethod[i].outpar = '<li class="cSquare"><i>Out Parameters:</i><li class="cCircle">'+ oOutParam.name + " As " + oOutParam.type1 + " <" + oOutParam.cimtype + ">";
			aMethod[i].inpar = sDescription = "";
			
			// InParameter
			if(aInParam = wmi_meth_inparameter(oItem)){
				for(var sInpar = "", j = 0, len = aInParam.length; j < len; j++){
					sDesc = (aInParam[j].desc != "") ? "<br>(" + aInParam[j].desc + ")": "";
					sValue = (aInParam[j].value != "") ? ", " + aInParam[j].value : "";
					sInpar += '<li class="cCircle"><i>' + aInParam[j].name + "</i> As " + aInParam[j].type1 + " <" + aInParam[j].cimtype + ">" + sDesc + sValue;
				}
				aMethod[i].inpar = '<li class="cSquare"><i>In Parameters:</i>' + sInpar;
				sDescription = (aInParam.desc == "") ? "&nbsp;" : aInParam.desc;
			}
			
			// Qualifier
			sQualifier = wmi_obj_qualifier(oItem,'<li class="cCircle">');
			
			aMethod.table += '\n<tr valign="top"><td>' + (i+1) + '&nbsp;</td><td><b>' + aMethod[i].name + "</b>()" + aMethod[i].inpar + aMethod[i].outpar + '</td><td>' + sDescription + '</td></tr>';
			aMethod.table += '\n<tr><td>&nbsp;</td><td colspan="2"><table width="100%" bgcolor="#EEEEEE" border="0" cellspacing="0" cellpadding="0">';
			aMethod.table += '\n\t<tr><td><li class="cSquare"><i>Qualifiers:</i><br>' + (sQualifier ? sQualifier : "") + '</td></tr>';
			aMethod.table += '</table></td></tr>';
		}
		aMethod.table += '\n</table>';
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return aMethod;
	}
}

function wmi_class_property(sOpt,oService,sClass){
	var oEnum
	try{ // http://www.microsoft.com/technet/scriptcenter/scrguide/sas_wmi_yzik.asp
		oService = oService ? oService : wmi_wbem_service();
		var aProperty = new Array();
		var oClass = oService.Get(sClass);
		
		if(!(oEnum = new Enumerator(oClass.Properties_))) return false
		else if(sOpt == "gethtmltext"){
			var p = '<table width="100%" border="0" cellspacing="0" cellpadding="0" class="cTable1">';
			p = p + '\n<caption align="left"><i>Properties for ' + sClass + '</i></caption>';
			p = p + '\n<tr align="left"><th width="10">&nbsp;</th><th>Name</th><th>Type</th><th>Array</th><th>Description</th></tr>';
			
			for(var oItem, i = 0; !oEnum.atEnd(); i++, oEnum.moveNext()){
				oItem = oEnum.item();
				aProperty[i] = new Object();
				aProperty[i].name = oItem.Name;
				// Qualifiers
				sQualifier = wmi_obj_qualifier(oItem,'<li class="cCircle">');
				aProperty[i].type = (wmi_class_parametertype(oItem)).type2;
				sDescription = (tmp = wmi_obj_description(oItem)) ? tmp : "&nbsp;";
				
				p = p + '\n<tr valign="top"><td width="12">' + (i+1) + '&nbsp;</td><td><b>' + oItem.Name + (oItem.IsArray ? "</b>[]" : "</b>") + '</td><td>' + aProperty[i].type +'</td><td>' + oItem.IsArray + '</td><td>' + sDescription + '</td></tr>';
				p = p + '\n<tr><td>&nbsp;</td><td colspan="4"><table width="100%" bgcolor="#EEEEEE" border="0" cellspacing="0" cellpadding="0">';
				p = p + '\n\t<tr><td><li class="cSquare"><i>Qualifiers:</i>' + sQualifier + '</td></tr>';
				p = p + '</table></td></tr>';
			}
			p = p + '\n</table>';
			aProperty.table = p
		}
		else if(sOpt == "getobject"){
			for(var oItem, o, n1, n2; !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				o = new Object()
				o.name = n1 = oItem.name
				n2 = n1.substring(1,n1.length), n2 = n2.replace(/[_]/g,"")
				o.name2 = n1.substring(0,1) + n2.replace(/([A-Z_])/g," $1") // Create spaces for every big letters except the first one
				o.description = (tmp = wmi_obj_description(oItem)) ? tmp : "";
				o.isarray = oItem.IsArray ? true : false;
				o.isdatetime = oItem.CIMType == 101 ? true : false;
				o.type = oItem.CIMType
				aProperty.push(o)
			}
		}
		else if(sOpt == "getobject_simple"){
			for(var oItem, o, n1, n2; !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				o = new Object()
				o.name = n1 = oItem.name
				//n2 = n1.substring(1,n1.length), n2 = n2.replace(/[_]/g,"")
				//o.name2 = n1.substring(0,1) + n2.replace(/([A-Z_])/g," $1") // Create spaces for every big letters except the first one
				//o.description = ""
				o.isarray = oItem.IsArray ? true : false;
				o.isdatetime = oItem.CIMType == 101 ? true : false;
				//o.type = oItem.CIMType
				aProperty.push(o)
			}
		}
		return aProperty;
	}
	catch(e){
		var d = js_str_trim(e.description)
		if(!d.match(/not found/ig)) wmi_log_error(2,e);
		return false;
	}
	finally{
		js_str_kill(oEnum)
	}
}

function wmi_class_qualifier(sClass,oService){
	try{
		oService = oService ? oService : wmi_wbem_service();
		var aQualifier = new Array();
		var oClass = oService.Get(sClass);
		var q = '<table border="0" cellspacing="1" cellpadding="0" class="cTableA">';
		q = q + '\n<caption align="left"><i>Qualifiers for ' + sClass + '</i></caption>';
		q = q + '<tr><th width="10">&nbsp;</th><th>Name</th><th>Value</th><th>&nbsp;</th></tr>';
		for(var oItem = "", i = 0, oEnum = new Enumerator(oClass.Qualifiers_); !oEnum.atEnd(); i++, oEnum.moveNext()){
			oItem = oEnum.item();			
			aQualifier[i] = new Object();
			aQualifier[i].name = oItem.Name;
			sValue = (oItem.Name.toLowerCase() == "valuemap") ? oItem.Value.replace(/,+/g,", ") : oItem.Value; // To long string, must break
			aQualifier[i].value = sValue;
			aQualifier[i].text = oItem.Name + " = " + sValue;
			q = q + '<tr><td>' + (i+1) + '</td><td>' + oItem.Name + '</td><td>' + sValue + '</td><td>&nbsp;</td></tr>';
		}
		oEnum = null
		q = q + '</table>';
		aQualifier.table = q
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return aQualifier;
	}
}

function wmi_class_mof(sClass,oService){
	try{ // Obtains the textual rendition of the object in Managed Object Format (MOF) syntax
		oService = oService ? oService : wmi_wbem_service();
		var oMOF = oService.Get(sClass,wmi.FlagUseAmendedQualifiers);
		var sMOF = oMOF.GetObjectText_();
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sMOF;
	}
}

function wmi_class_path(sClass,oService){
	try{
		var sPath = false;
		oService = oService ? oService : wmi_wbem_service();
		try{
			var oClass = oService.Get(sClass);
			sPath = oClass.Path_;
		}
		catch(ee){
			// Needs to add WMI error object
			// This will error: "Not Found" or "Nothing")
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sPath;
	}
}

function wmi_class_parametertype(oProperty){
	try{
	/*
	  CIM_ILLEGAL CIM_ILLEGAL An illegal value. = 4095,    // 0xFFF
  CIM_EMPTY CIM_EMPTY An empty (null) value. = 0,    // 0x0
  CIM_SINT8 CIM_SINT8 An 8-bit signed integer. = 16,    // 0x10
  CIM_UINT8 CIM_UINT8 An 8-bit unsigned integer. = 17,    // 0x11
  CIM_SINT16 CIM_SINT16 A 16-bit signed integer. = 2,    // 0x2
  CIM_UINT16 CIM_UINT16 A 16-bit unsigned integer. = 18,    // 0x12
  CIM_SINT32 CIM_SINT32 A 32-bit signed integer. = 3,    // 0x3
  CIM_UINT32 CIM_UINT32 A 32-bit unsigned integer. = 19,    // 0x13
  CIM_SINT64 CIM_SINT64 A 64-bit signed integer. = 20,    // 0x14
  CIM_UINT64 CIM_UINT64 A 64-bit unsigned integer. = 21,    // 0x15
  CIM_REAL32 CIM_REAL32 A 32-bit real number. = 4,    // 0x4
  CIM_REAL64 CIM_REAL64 A 64-bit real number. = 5,    // 0x5
  CIM_BOOLEAN CIM_BOOLEAN A Boolean value. = 11,    // 0xB
  CIM_STRING CIM_STRING A string value. = 8,    // 0x8
  CIM_DATETIME CIM_DATETIME A DateTime value. = 101,    // 0x65
  CIM_REFERENCE CIM_REFERENCE Reference (__Path) of another Object. = 102,    // 0x66
  CIM_CHAR16 CIM_CHAR16 A 16-bit character value. = 103,    // 0x67
  CIM_OBJECT CIM_OBJECT An Object value. = 13,    // 0xD
  CIM_FLAG_ARRAY CIM_FLAG_ARRAY An array value. = 8192    // 0x2000

	
	*/
	
		var oType = new Object();		
		switch(oProperty.CIMType){
			case 2 : {
				oType.type1 = "Integer", oType.type2 = "SINT16";
				break;
			}
			case 3 : {
				oType.type1 = "Integer", oType.type2 = "SINT32";
				break;
			}
			case 4 : {
				oType.type1 = "Integer", oType.type2 = "REAL32";
				break;
			}
			case 5 : {
				oType.type1 = "Integer", oType.type2 = "REAL64";
				break;
			}
			case 8 : {
				oType.type1 = "String", oType.type2 = "BOOLEAN";
				break;
			}
			case 11 : {
				oType.type1 = "Boolean", oType.type2 = "";
				break;
			}
			case 13 : {
				oType.type1 = "Object", oType.type2 = "CIM_OBJECT";
				break;
			}
			case 16 : {
				oType.type1 = "Integer", oType.type2 = "SINT8";
				break;
			}
			case 17 : {
				oType.type1 = "Integer", oType.type2 = "USINT8";
				break;
			}
			case 18 : {
				oType.type1 = "Integer", oType.type2 = "USINT16";
				break;
			}
			case 19 : {
				oType.type1 = "Integer", oType.type2 = "USINT32";
				break;
			}
			case 20 : {
				oType.type1 = "Integer", oType.type2 = "SINT64";
				break;
			}
			case 21 : {
				oType.type1 = "Integer", oType.type2 = "USINT64";
				break;
			}
			case 101 : {
				oType.type1 = "Datetime", oType.type2 = "";
				break;
			}
			case 102 : { // Reference to a CIM object.
				oType.type1 = "Reference", oType.type2 = "REF";
				break;
			}
			case 103 : {
				oType.type1 = "Char 16", oType.type2 = "UNICODE";
				break;
			}
			default : oType.type1 = "Unknown", oType.type2 = ""; break;
		}
		if(oProperty.IsArray) oType.type1 += "[]";
		oType.text = oType.type1 + (oType.type2 == "" ? "" : " (" + oType.type2 + ")")
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return oType;
	}
}

function wmi_obj_description(oQualObject){
	try{
		var sDescription = "";
		var oQualifier = oQualObject.Qualifiers_;
		for(var oItem, oEnum = new Enumerator(oQualObject.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			if(oItem.Name.toLowerCase() == "description"){
				try{ sDescription = oItem.Value; }
				catch(ee){};
				break;
			}
		}
		oEnum = null
	}	
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sDescription;
	}
}

function wmi_obj_qualifier(oQual,sStr){
	try{ // http://msdn.microsoft.com/library/en-us/wmisdk/wmi/standard_qualifiers.asp
		var sValue = "", aContext;
		for(var oItem, i = 1, oEnum = new Enumerator(oQual.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item();
			var sStr2 = sStr ? sStr : i + ". ";
			if(aContext = wmi_obj_isvbarray(oItem.Value)){
				if(oItem.Name.toLowerCase() == "valuemap"){
					aContext = (aContext.toString()).replace(/,+/ig,", "); // To long string, must break
				}
				sValue = sValue + sStr2 + oItem.Name + " = " + aContext + "\n";
			}
			else sValue = sValue + sStr2 + oItem.Name + " = " + oItem.Value + "\n";
		}
	}	
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sValue.replace(/(.*)\n$/g,"$1");
	}
}

function wmi_obj_qualifiername(oQual){
	try{ // http://msdn.microsoft.com/library/en-us/wmisdk/wmi/standard_qualifiers.asp
		var dQual = new ActiveXObject("Scripting.Dictionary");
		for(var oItem, oEnum = new Enumerator(oQual.Qualifiers_); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			dQual.Add(oItem.Name,"")
		}
	}	
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return dQual;
	}
}

function wmi_meth_inparameter(oMethod){
	try{
		var oInpars = oMethod.InParameters; // EX: Method parameters for Win32_Bus
		if(typeof(oInpars) != null){ // Some method parameters are null
			try{ // Properties may not exist
				var a = new Array(), tmp;
				for(var oType, i = 0, oEnum = new Enumerator(oInpars.Properties_); !oEnum.atEnd(); i++, oEnum.moveNext()){
					var oItem = oEnum.item(); // EX: SetPowerState
					a[i] = new Object(); // EX: PowerState 
					a[i].name = a[i].type1 = a[i].typestr = a[i].value = "";
					oType = wmi_class_parametertype(oItem)
					a[i].name = oItem.Name // EX: PowerState  
					a[i].type1 = oType.type1; // EX: Integer
					a[i].typestr = oType.text; //EX: Integer (USINT16)
					a[i].cimtype = oItem.CIMtype;
					// Get the description
					a[i].desc = (tmp = wmi_obj_description(oItem)) ? tmp : "";					
				}
			}
			catch(ee){ }
			a.name = oMethod.Name;
			a.desc = wmi_obj_description(oMethod); // Get description			
		}
		return a;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		js_str_kill(oInpars)
	}
}

function wmi_meth_outparameter(oMethod){
	try{
		var oOutpars = oMethod.OutParameters; // EX: Method parameters for Win32_Bus
		if(typeof(oOutpars) != null){ // Some method parameters are null
			try{ // Properties may not exist
				var o = new Object();
				for(var i = 0, oEnum = new Enumerator(oOutpars.Properties_); !oEnum.atEnd(); i++, oEnum.moveNext()){
					var oItem = oEnum.item(); // EX: SetPowerState
					o.name = (oItem.Name)
					o.type1 = (wmi_class_parametertype(oItem)).type1;
					o.typestr = (wmi_class_parametertype(oItem)).text;
					o.value = oItem.Value;
					o.cimtype = oItem.CIMtype;		
				}
				return o
			}
			catch(ee){}			
		}
		return false;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{		
		js_str_kill(oOutpars)
	}
}

function wmi_class_echo(aClass){
	try{
		for(var i in aClass){
			for(var j = 0, len = aClass[i].length; j < len; j++){				
				WScript.Echo(((i==0)?"":"\t") + aClass[i][j]);
			}
		}
		return true;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_obj_isvbarray(aVBArray){
	try{
		var aJSArray = false;
		if(typeof(aVBArray) == "unknown"){
			aJSArray = new VBArray(aVBArray).toArray();
		}
		return aJSArray;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_class_privilege(sClass){
	try{
		/*
wbemPrivilegeCreateToken
1 Required to create a primary token. 
wbemPrivilegePrimaryToken
2 Required to assign the primary token of a process. 
wbemPrivilegeLockMemory
3 Required to lock physical pages in memory. 
wbemPrivilegeIncreaseQuota
4 Required to increase the quota assigned to a process. 
wbemPrivilegeMachineAccount
5 Required to create a machine account. 
wbemPrivilegeTcb
6 Identifies its holder as part of the trusted computer base. Some trusted, protected subsystems are granted this privilege. 
wbemPrivilegeSecurity
7 Required to perform a number of security-related functions, such as controlling and viewing audit messages. This privilege identifies its holder as a security operator. 
wbemPrivilegeTakeOwnership
8 Required to take ownership of an object without being granted discretionary access. This privilege allows the owner value to be set only to those values that the holder may legitimately assign as the owner of an object. 
wbemPrivilegeLoadDriver
9 Required to load or unload a device driver. 
wbemPrivilegeSystemProfile
10 Required to gather profiling information for the entire system. 
wbemPrivilegeSystemtime
11 Required to modify the system time. 
wbemPrivilegeProfileSingleProcess
12 Required to gather profiling information for a single process. 
wbemPrivilegeIncreaseBasePriority
13 Required to increase the base priority of a process. 
wbemPrivilegeCreatePagefile
14 Required to create a paging file. 
wbemPrivilegeCreatePermanent
15 Required to create a permanent object. 
wbemPrivilegeBackup
16 Required to perform backup operations. 
wbemPrivilegeRestore
17 Required to perform restore operations. This privilege enables you to set any valid user or group security identifier (SID) as the owner of an object. 
wbemPrivilegeShutdown
18 Required to shut down a local system. 
wbemPrivilegeDebug
19 Required to debug a process. 
wbemPrivilegeAudit
20 Required to generate audit-log entries. 
wbemPrivilegeSystemEnvironment
21 Required to modify the nonvolatile RAM of systems that use this type of memory to store configuration information. 
wbemPrivilegeChangeNotify
22 Required to receive notifications of changes to files or directories. This privilege also causes the system to skip all traversal access checks. It is enabled by default for all users. 
wbemPrivilegeRemoteShutdown
23 Required to shut down a system using a network request. 
wbemPrivilegeUndock
24 Required to remove computer from docking station. 
wbemPrivilegeSyncAgent
25 Required to synchronize directory service data. 
wbemPrivilegeEnableDelegation
26 Required to enable computer and user accounts to be trusted for delegation. 

		
		
		*/
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

/// Getting Classses

function wmi_win32_ipaddress(oService,iIndex,bDhcp){
	try{
		oService = oService ? oService : wmi_wbem_service();
		sIndex = js_str_isnumber(iIndex) ? " Index=" + iIndex : "";
		sDhcp = bDhcp ? " DHCPEnabled='true'" : " DHCPEnabled='false'";
		var oColItems = oService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' AND " + sIndex + sDhcp,"WQL",48);
		for(var oItem, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			try{
				v = wmi_class_vbarray(oItem.IPAddress,'IPAddress');
				a = (v.stream).split("  ");
				return a[2];
			}
			catch(ee){}
		}
		return false;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_win32_bios(oService){
	try{
		oService = oService ? oService : wmi_wbem_service();
		var sBios = "";
		var oBios = oService.InstancesOf("Win32_Bios");
		for(var oItem, oEnum = new Enumerator(oBios); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			sBios += "\n\nComputer:\t\t" + oWno.ComputerName;
			sBios += "\nBios Name:\t" + oItem.Name;
			sBios += "\nBuild Number:\t" + oItem.BuildNumber;
			sBios += "\nRelease Date:\t" + (d=wmi_class_date(oItem.ReleaseDate)) ? d : "";
			sBios += "\nManufacturer:\t" + oItem.Manufacturer;
			sBios += "\nSerial Number:\t" + oItem.SerialNumber;
			sBios += "\nSMB Version:\t" + oItem.SMBIOSBIOSVersion;
			sBios += "\nBios Version:\t" + oItem.Version;
		}
		
		//
		oBios = oService.InstancesOf("Win32_SystemEnclosure");
		for(var oItem, oEnum = new Enumerator(oBios); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			sBios += "\nPart Number:\t" + oItem.PartNumber;
			sBios += "\nAsset Tag:\t" + oItem.SMBIOSAssetTag;
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sBios;
	}
}

function wmi_win32_pagefile(sComputer,sNameSpace,sImpersonation,sAuthentication){
	try{
		var sHtml, sHtmlAll = "";
		sComputer = sComputer ? sComputer : ".";
		sNameSpace = sNameSpace ? sNameSpace : "root\\CIMV2";
		sImpersonation = sImpersonation ? sImpersonation : "Impersonate";
		sAuthentication = sAuthentication ? sAuthentication : "Connect";
		var oService = GetObject("winmgmts:{impersonationLevel=" + sImpersonation + ",authenticationLevel=" + sAuthentication + ",(Security)}!\\\\" + sComputer + "\\" + sNameSpace);
		var oColItems = oService.ExecQuery("Select * from Win32_PageFile","WQL",48);

		for(var oItem, i = 1, d, v, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item(), sHtml = "\n [INSTANCE-" + i + "]";
			sHtml = sHtml + "\n  AccessMask: " + oItem.AccessMask;
			sHtml = sHtml + "\n  Archive: " + oItem.Archive;
			sHtml = sHtml + "\n  Caption: " + oItem.Caption;
			sHtml = sHtml + "\n  Compressed: " + oItem.Compressed;
			sHtml = sHtml + "\n  CompressionMethod: " + oItem.CompressionMethod;
			sHtml = sHtml + "\n  CreationClassName: " + oItem.CreationClassName;
			sHtml = sHtml + "\n  CreationDate: " + ((d=wmi_class_date(oItem.CreationDate)) ? d : "undefined");
			sHtml = sHtml + "\n  CSCreationClassName: " + oItem.CSCreationClassName;
			sHtml = sHtml + "\n  CSName: " + oItem.CSName;
			sHtml = sHtml + "\n  Description: " + oItem.Description;
			sHtml = sHtml + "\n  Drive: " + oItem.Drive;
			sHtml = sHtml + "\n  EightDotThreeFileName: " + oItem.EightDotThreeFileName;
			sHtml = sHtml + "\n  Encrypted: " + oItem.Encrypted;
			sHtml = sHtml + "\n  EncryptionMethod: " + oItem.EncryptionMethod;
			sHtml = sHtml + "\n  Extension: " + oItem.Extension;
			sHtml = sHtml + "\n  FileName: " + oItem.FileName;
			sHtml = sHtml + "\n  FileSize: " + oItem.FileSize;
			sHtml = sHtml + "\n  FileType: " + oItem.FileType;
			sHtml = sHtml + "\n  FreeSpace: " + oItem.FreeSpace;
			sHtml = sHtml + "\n  FSCreationClassName: " + oItem.FSCreationClassName;
			sHtml = sHtml + "\n  FSName: " + oItem.FSName;
			sHtml = sHtml + "\n  Hidden: " + oItem.Hidden;
			sHtml = sHtml + "\n  InitialSize: " + oItem.InitialSize;
			sHtml = sHtml + "\n  InstallDate: " + ((d=wmi_class_date(oItem.InstallDate)) ? d : "undefined");
			sHtml = sHtml + "\n  InUseCount: " + oItem.InUseCount;
			sHtml = sHtml + "\n  LastAccessed: " + ((d=wmi_class_date(oItem.LastAccessed)) ? d : "undefined");
			sHtml = sHtml + "\n  LastModified: " + ((d=wmi_class_date(oItem.LastModified)) ? d : "undefined");
			sHtml = sHtml + "\n  Manufacturer: " + oItem.Manufacturer;
			sHtml = sHtml + "\n  MaximumSize: " + oItem.MaximumSize;
			sHtml = sHtml + "\n  Name: " + oItem.Name;
			sHtml = sHtml + "\n  Path: " + oItem.Path;
			sHtml = sHtml + "\n  Readable: " + oItem.Readable;
			sHtml = sHtml + "\n  Status: " + oItem.Status;
			sHtml = sHtml + "\n  System: " + oItem.System;
			sHtml = sHtml + "\n  Version: " + oItem.Version;
			sHtml = sHtml + "\n  Writeable: " + oItem.Writeable;
			sHtmlAll = sHtmlAll + sHtml;
			WScript.Echo(sHtml);
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sHtmlAll;
	}
}

function wmi_win32_registry(sComputer,sNameSpace,sImpersonation,sAuthentication){
	try{
		var sHtml, sHtmlAll = "";
		sComputer = sComputer ? sComputer : ".";
		sNameSpace = sNameSpace ? sNameSpace : "root\\CIMV2";
		sImpersonation = sImpersonation ? sImpersonation : "Impersonate";
		sAuthentication = sAuthentication ? sAuthentication : "Connect";
		var oService = GetObject("winmgmts:{impersonationLevel=" + sImpersonation + ",authenticationLevel=" + sAuthentication + ",(Security)}!\\\\" + sComputer + "\\" + sNameSpace);
		var oColItems = oService.ExecQuery("Select * from Win32_Registry","WQL",48);

		for(var oItem, i = 1, d, v, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item(), sHtml = "\n [INSTANCE-" + i + "]";
			sHtml = sHtml + "\n  Caption: " + oItem.Caption;
			sHtml = sHtml + "\n  CurrentSize: " + oItem.CurrentSize;
			sHtml = sHtml + "\n  Description: " + oItem.Description;
			sHtml = sHtml + "\n  InstallDate: " + ((d=wmi_class_date(oItem.InstallDate)) ? d : "undefined");
			sHtml = sHtml + "\n  MaximumSize: " + oItem.MaximumSize;
			sHtml = sHtml + "\n  Name: " + oItem.Name;
			sHtml = sHtml + "\n  ProposedSize: " + oItem.ProposedSize;
			sHtml = sHtml + "\n  Status: " + oItem.Status;
			sHtmlAll += sHtml;
			WScript.Echo(sHtml);
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sHtmlAll;
	}
}

function wmi_win32_quickfixengineering(sComputer,sNameSpace,sImpersonation,sAuthentication){
	try{
		var sHtml, sHtmlAll = "";
		sComputer = sComputer ? sComputer : ".";
		sNameSpace = sNameSpace ? sNameSpace : "root\\CIMV2";
		sImpersonation = sImpersonation ? sImpersonation : "Impersonate";
		sAuthentication = sAuthentication ? sAuthentication : "Connect";
		var oService = GetObject("winmgmts:{impersonationLevel=" + sImpersonation + ",authenticationLevel=" + sAuthentication + ",(Security)}!\\\\" + sComputer + "\\" + sNameSpace);
		var oColItems = oService.ExecQuery("Select * from Win32_QuickFixEngineering","WQL",48);

		for(var oItem, i = 1, d, v, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item(), sHtml = "\n [INSTANCE-" + i + "]";
			sHtml = sHtml + "\n  Caption: " + oItem.Caption;
			sHtml = sHtml + "\n  CSName: " + oItem.CSName;
			sHtml = sHtml + "\n  Description: " + oItem.Description;
			sHtml = sHtml + "\n  FixComments: " + oItem.FixComments;
			sHtml = sHtml + "\n  HotFixID: " + oItem.HotFixID;
			sHtml = sHtml + "\n  InstallDate: " + ((d=wmi_class_date(oItem.InstallDate)) ? d : "undefined");
			sHtml = sHtml + "\n  InstalledBy: " + oItem.InstalledBy;
			sHtml = sHtml + "\n  InstalledOn: " + oItem.InstalledOn;
			sHtml = sHtml + "\n  Name: " + oItem.Name;
			sHtml = sHtml + "\n  ServicePackInEffect: " + oItem.ServicePackInEffect;
			sHtml = sHtml + "\n  Status: " + oItem.Status;
			sHtmlAll = sHtmlAll + sHtml;
			WScript.Echo(sHtml);
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sHtmlAll;
	}
}

function wmi_win32_networkadapterconfiguration(bOpt1,bOpt2,oService,sComputer,sUser,sPass,sImpersonation,sAuthentication){
	try{
		var sHtml, sHtmlAll = "";
		var oObject, aObject = new Array()
		var sQuery = bOpt2 ? " and DHCPEnabled='true'" : ""
		var oService = oService ? oService : wmi_wbem_service(sComputer,"root\\cimv2",sUser,sPass,sImpersonation,null,sAuthentication);
		var oColItems = oService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled='true'" + sQuery,"WQL",48);
		for(var oItem, i = 1, d, v, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext(), i++){
			oItem = oEnum.item(), sHtml = "\n [INSTANCE-" + i + "]";
			oObject = new Object()
			sHtml = sHtml + "\n  ArpAlwaysSourceRoute: " + oItem.ArpAlwaysSourceRoute;
			sHtml = sHtml + "\n  ArpUseEtherSNAP: " + oItem.ArpUseEtherSNAP;
			sHtml = sHtml + "\n  Caption: " + oItem.Caption;
			sHtml = sHtml + "\n  DatabasePath: " + oItem.DatabasePath;
			sHtml = sHtml + "\n  DeadGWDetectEnabled: " + oItem.DeadGWDetectEnabled;
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.DefaultIPGateway,'DefaultIPGateway')) ? v.stream : "\n  DefaultIPGateway: " + "undefined");
			oObject.IPGateway = v ? v.array[0] : "";
			sHtml = sHtml + "\n  DefaultTOS: " + oItem.DefaultTOS;
			sHtml = sHtml + "\n  DefaultTTL: " + oItem.DefaultTTL;
			sHtml = sHtml + "\n  Description: " + (oObject.Description = oItem.Description)
			sHtml = sHtml + "\n  DHCPEnabled: " + (oObject.DHCPEnabled = oItem.DHCPEnabled)
			sHtml = sHtml + "\n  DHCPLeaseExpires: " + ((d1=wmi_class_date(oItem.DHCPLeaseExpires)) ? d1 : "undefined");
			sHtml = sHtml + "\n  DHCPLeaseObtained: " + ((d2=wmi_class_date(oItem.DHCPLeaseObtained)) ? d2 : "undefined");
			sHtml = sHtml + "\n  DHCPServer: " + oItem.DHCPServer
			if(oItem.DHCPEnabled){
				oObject.DHCPLeaseExpires = d1
				oObject.DHCPLeaseObtained = d2
				oObject.DHCPServer = oItem.DHCPServer
			}
			sHtml = sHtml + "\n  DNSDomain: " + (oObject.DNSDomain = oItem.DNSDomain)
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.DNSDomainSuffixSearchOrder,'DNSDomainSuffixSearchOrder')) ? v.stream : "\n  DNSDomainSuffixSearchOrder: " + "undefined");
			oObject.DNSDomainSuffixSearchOrder = v ? (v.array).toString() : ""
			sHtml = sHtml + "\n  DNSEnabledForWINSResolution: " + oItem.DNSEnabledForWINSResolution;
			sHtml = sHtml + "\n  DNSHostName: " + (oObject.DNSHostName = oItem.DNSHostName)
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.DNSServerSearchOrder,'DNSServerSearchOrder')) ? v.stream : "\n  DNSServerSearchOrder: " + "undefined");
			oObject.DNSServerSearchOrder = v ? (v.array).toString() : ""
			sHtml = sHtml + "\n  DomainDNSRegistrationEnabled: " + (oObject.DomainDNSRegistrationEnabled = oItem.DomainDNSRegistrationEnabled)
			sHtml = sHtml + "\n  ForwardBufferMemory: " + oItem.ForwardBufferMemory;
			sHtml = sHtml + "\n  FullDNSRegistrationEnabled: " + (oObject.FullDNSRegistrationEnabled = oItem.FullDNSRegistrationEnabled);
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.GatewayCostMetric,'GatewayCostMetric')) ? v.stream : "\n  GatewayCostMetric: " + "undefined");
			sHtml = sHtml + "\n  IGMPLevel: " + oItem.IGMPLevel;
			sHtml = sHtml + "\n  Index: " + oItem.Index;
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPAddress,'IPAddress')) ? v.stream : "\n  IPAddress: " + "undefined");
			oObject.IPAddress = v ? (v.array).toString() : ""
			sHtml = sHtml + "\n  IPConnectionMetric: " + oItem.IPConnectionMetric;
			sHtml = sHtml + "\n  IPEnabled: " + oItem.IPEnabled;
			sHtml = sHtml + "\n  IPFilterSecurityEnabled: " + oItem.IPFilterSecurityEnabled;
			sHtml = sHtml + "\n  IPPortSecurityEnabled: " + oItem.IPPortSecurityEnabled;
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPSecPermitIPProtocols,'IPSecPermitIPProtocols')) ? v.stream : "\n  IPSecPermitIPProtocols: " + "undefined");
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPSecPermitTCPPorts,'IPSecPermitTCPPorts')) ? v.stream : "\n  IPSecPermitTCPPorts: " + "undefined");
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPSecPermitUDPPorts,'IPSecPermitUDPPorts')) ? v.stream : "\n  IPSecPermitUDPPorts: " + "undefined");
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPSubnet,'IPSubnet')) ? v.stream : "\n  IPSubnet: " + "undefined");
			oObject.IPSubnet = v ? (v.array).toString() : ""
			sHtml = sHtml + "\n  IPUseZeroBroadcast: " + oItem.IPUseZeroBroadcast;
			sHtml = sHtml + "\n  IPXAddress: " + oItem.IPXAddress;
			sHtml = sHtml + "\n  IPXEnabled: " + oItem.IPXEnabled;
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPXFrameType,'IPXFrameType')) ? v.stream : "\n  IPXFrameType: " + "undefined");
			sHtml = sHtml + "\n  IPXMediaType: " + oItem.IPXMediaType;
			sHtml = sHtml + ((v=wmi_class_vbarray(oItem.IPXNetworkNumber,'IPXNetworkNumber')) ? v.stream : "\n  IPXNetworkNumber: " + "undefined");
			sHtml = sHtml + "\n  IPXVirtualNetNumber: " + oItem.IPXVirtualNetNumber;
			sHtml = sHtml + "\n  KeepAliveInterval: " + oItem.KeepAliveInterval;
			sHtml = sHtml + "\n  KeepAliveTime: " + oItem.KeepAliveTime;
			sHtml = sHtml + "\n  MACAddress: " + (oObject.MACAddress = oItem.MACAddress)
			sHtml = sHtml + "\n  MTU: " + oItem.MTU;
			sHtml = sHtml + "\n  NumForwardPackets: " + oItem.NumForwardPackets;
			sHtml = sHtml + "\n  PMTUBHDetectEnabled: " + oItem.PMTUBHDetectEnabled;
			sHtml = sHtml + "\n  PMTUDiscoveryEnabled: " + oItem.PMTUDiscoveryEnabled;
			sHtml = sHtml + "\n  ServiceName: " + (oObject.ServiceName = oItem.ServiceName)
			sHtml = sHtml + "\n  SettingID: " + oItem.SettingID;
			sHtml = sHtml + "\n  TcpipNetbiosOptions: " + oItem.TcpipNetbiosOptions;
			sHtml = sHtml + "\n  TcpMaxConnectRetransmissions: " + oItem.TcpMaxConnectRetransmissions;
			sHtml = sHtml + "\n  TcpMaxDataRetransmissions: " + oItem.TcpMaxDataRetransmissions;
			sHtml = sHtml + "\n  TcpNumConnections: " + oItem.TcpNumConnections;
			sHtml = sHtml + "\n  TcpUseRFC1122UrgentPointer: " + oItem.TcpUseRFC1122UrgentPointer;
			sHtml = sHtml + "\n  TcpWindowSize: " + oItem.TcpWindowSize;
			sHtml = sHtml + "\n  WINSEnableLMHostsLookup: " + oItem.WINSEnableLMHostsLookup;
			sHtml = sHtml + "\n  WINSHostLookupFile: " + oItem.WINSHostLookupFile;
			sHtml = sHtml + "\n  WINSPrimaryServer: " + (oObject.WINSPrimaryServer = oItem.WINSPrimaryServer)
			sHtml = sHtml + "\n  WINSScopeID: " + oItem.WINSScopeID;
			sHtml = sHtml + "\n  WINSSecondaryServer: " + (oObject.WINSSecondaryServer = oItem.WINSSecondaryServer)
			
			if(bOpt1) sHtmlAll = sHtmlAll + sHtml;
			else aObject.push(oObject);
		}
		return bOpt1 ? sHtmlAll : aObject
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_win32_reboot(oService,sComputer){
	try{
		oService = oService ? oService :  GetObject("winmgnt:{impersonationLevel=impersonate},(Shutdown),(RemoteShutdown)}\\\\" + sComputer);
		var oColItems = oService.ExecQuery("Select * from Win32_OperatingSystem where Primary=true","WQL",48);
		for(var oItem, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			oItem.Reboot();
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_win32_computersystem_domainrole(oService,sComputer){
	try{
		oService = oService ? oService :  wmi_wbem_service(sComputer);
		var oColItems = oService.ExecQuery("Select DomainRole from Win32_ComputerSystem","WQL",48);
		for(var oItem, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			var o = new Object()
			o.number = oItem.DomainRole, o.name = "DomainRole"
			switch(parseInt(o.number)){
				case 0 : o.type = "Standalone Workstation"; break;
				case 1 : o.type = "Member Workstation"; break;
				case 2 : o.type = "Standalone Server"; break;
				case 3 : o.type = "Member Server"; break;
				case 4 : o.type = "Backup Domain Controller"; break;
				case 5 : o.type = "Primary Domain Controller"; break;
				default : o.type = "Unknown"; break;
			}
			return o
		}
		return false
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_win32_systemenclosure_chassistype(oService,sComputer){
	try{
		oService = oService ? oService :  wmi_wbem_service(sComputer);
		var oColItems = oService.ExecQuery("Select ChassisTypes from Win32_SystemEnclosure","WQL",48);
		for(var oItem, oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			var o = new Object()
			var v = new VBArray(oItem.ChassisTypes).toArray(), v = v[0]
			o.number = v, o.name = "ChassisType"
			switch(parseInt(o.number)){
				case 1 : o.type = "Other"; break;
				case 2 : o.type = "Unknown"; break;
				case 3 : o.type = "Desktop"; break;
				case 4 : o.type = "Low-profile Desktop"; break;
				case 5 : o.type = "Pizza Box"; break;
				case 6 : o.type = "Mini Tower"; break;
				case 7 : o.type = "Tower"; break;
				case 8 : o.type = "Portable"; break;
				case 9 : o.type = "Laptop"; break;
				case 10 : o.type = "Notebook"; break;
				case 11 : o.type = "Hand-held"; break;
				case 12 : o.type = "Docking Station"; break;
				case 13 : o.type = "All-in-one"; break;
				case 14 : o.type = "Sub Notebook"; break;
				case 15 : o.type = "Space-saving"; break;
				case 16 : o.type = "Lunch Box"; break;
				case 17 : o.type = "Main System Chassis"; break;
				case 18 : o.type = "Expansion Chassis"; break;
				case 19 : o.type = "Sub Chassis"; break;
				case 20 : o.type = "Bus-expansion Chassis"; break;
				case 21 : o.type = "Peripherical Chassis"; break;
				case 22 : o.type = "Storage Chassis"; break;
				case 23 : o.type = "Rack-mounted Chassis"; break;
				case 24 : o.type = "Sealed-case Computer"; break;
				default : o.type = "Unknown"; break;
			}
			return o
		}
		return false
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_win32_class(sOpt,sClass,oService,sQuery,sBreak,sComputer,sNameSpace,sUser,sPass,sImpersonation,sAuthentication,oTBody){
	var aProperties, sGetObject, bAsk = true, oSelect
	try{
		sQuery = sQuery ? sQuery : "Select * from " + sClass
		sNameSpace = sNameSpace ? sNameSpace : "root\\cimv2"
		oService = oService ? oService : wmi_wbem_service(sComputer,sNameSpace,sUser,sPass,sImpersonation,null,sAuthentication);		
		var oRe = new RegExp("Select[ \t]+(.+)[ \t]+from[ \t]+([a-z0-9_\-]+)[ \t]*(.*)","ig")
		var bSelect = false, oRe1
		if(!oService) return false
		if(sQuery.match(oRe)){
			oSelect = new ActiveXObject("Scripting.Dictionary");
			var a = (RegExp.$1).split(/[, \t]{1,}/ig)			
			if(a[0] != "*"){
				for(var i = 0, len = a.length; i < len; i++) oSelect.Add(a[i],a[i])
				if(sQuery.match(/ where (.*)/ig)){
					var t = RegExp.$1
					t = t.replace(/[ \t]{2,}/g," ")
					t = t.replace(/[ \t]*=[ \t]*/g,"=")
					t = t.replace(/[ \t]+and[ \t]+/ig," ")
					t = t.replace(/[ \t]+or[ \t]+/ig," ")
					t = t.replace(/[<>]+/g,"")
					t = t.replace(/["']+/g,"")
					t = t.replace(/[ ]/g,"=")
					a = t.split("=")
					for(var i = 0, len = a.length; i < len; i += 2){
						if(!oSelect.Exists(a[i])) oSelect.Add(a[i],a[i])
					}
				}
				bSelect = true
			}
			js_str_kill(a)
		}
		else return false
		var bTBody = typeof(oTBody) == "object" ? true : false // Can directly populate a table using TBody 		
		var bSimple = sOpt == "simple" ? true : false
		var bShow = sOpt == "show" ? true : false
		
		sGetObject = bSimple ? "getobject_simple" : "getobject"
		aProperties = wmi_class_property(sGetObject,oService,sClass)
		
		if(aProperties){
			sBreak = sBreak ? sBreak : "\n  "
			var oColItems = oService.ExecQuery(sQuery,"WQL",48), r = 0, rr;
			if(bSimple || bShow) var aClass = new Array()
			else if(bTBody){
				js_log_print("log_result","## Creating table for properties using query: " + sQuery)
				var r = oTBody.rows.length, rr, oRow, oCell
				var iTime = (new Date()).getTime()
			}
			var len = aProperties.length
			for(var i = 0, oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext(), i++){
				var oItem = oEnum.item();
				if(!bTBody){
					aClass[i] = new Array();
					if(bShow && i > 0) sClass = sClass + "\n"
				}
				else if(r > 0){ // do not create a row in the beginning 
					oRow = oTBody.insertRow((r++))
					oCell = oRow.insertCell(), oCell.innerText = "instance - " + js_str_number(i), oCell.className = "cHead2", oCell.colSpan = 5
				}
				
				for(var j = 0; j < len && !js.pro.stopit; j++, r++){
					var n = aProperties[j].name, v, t
					if(bSelect && !oSelect.Exists(n)) continue // Check if the query string has "*". If not, then exclude unchosen fields					
					try{
						if(aProperties[j].isarray){
							t = (v=wmi_class_vbarray(oItem[n],n,sBreak)) ? v.stream : " "
						}
						else if(aProperties[j].isdatetime) t = wmi_class_date(oItem[n])
						else t = oItem[n]
						if(bShow) sClass = sClass + "\n" + n + ": " + t
						else if(bTBody){
							oRow = oTBody.insertRow(r), oRow.style.verticalAlign = 'top', rr = js_str_number(r+1);
							oCell = oRow.insertCell(), oCell.innerText = rr, oCell.width = 22, oCell.align = "center"
							oCell = oRow.insertCell(), oCell.innerText = n, oCell.className = "cTable1TD", oCell.noWrap = true, oCell.width = "1%"
							oCell = oRow.insertCell(), oCell.innerText = t + " "
							oCell = oRow.insertCell()
							oCell.innerText = bSimple ? " " : aProperties[j].description							
							oCell = oRow.insertCell(), oCell.innerText = rr, oCell.width = 22, oCell.align = "center"
						}
						else {
							aClass[i][n] = new Object()
							aClass[i][n].name = n
							aClass[i][n].value = t
							if(!bSimple){
								aClass[i][n].description = aProperties[j].description
								aClass[i][n].string = n + ": " + t
							}
						}
					}
					catch(ee){
						//js_log_print("log_result","ERR: "+j + " "+n)
					}
					js_str_kill(n,v,t)
				}
				if(bTBody && bAsk){
					if(r < 1200 || ((new Date()).getTime()-iTime) < 120000) continue // contnue if less than 1200 rows or less than 2 minutes
					if(js_str_popup("The WQL query: " + sQuery + " is taking awfully a lot of time. Continue?",20000,"WMI enumeration popup",32 + 1) == 2) break;
					bAsk = false
				}
				//else if(!bShow) js_tme_sleep(js.time.mil3)
			}
		}		
		return (bTBody ? (r > 0) : aClass)
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		if(bSelect) oSelect.RemoveAll(), delete oSelect
		if(bShow) js_log_print("log_result",sClass)
		js_str_kill(aProperties)
	}
}	
			
function wmi_class_date(sDate){
	try{
		if(sDate != null){
			sDate = sDate.replace(/[\*]/g,0);
			var sDateTime = sDate.substring(0,4) + "-" + sDate.substring(4,6) + "-" + sDate.substring(6,8) + " " + sDate.substring(8,10) + ":" + sDate.substring(10,12) + ":" + sDate.substring(12,14) + "." + sDate.substring(15,18);
			return sDateTime;
		}
		else return false;
	}
	catch(e){
		return false;
	}
}

function wmi_class_vbarray(aVBarray,sVBName,sBreak){
	try{
		if(aVBarray != null){ // VBArray object
			var aVB = new VBArray(aVBarray).toArray();
			var aArray = new Array()
			sBreak = sBreak ? sBreak : "\n  "
			for(var j = 0, sVB = "", len = aVB.length; j < len; j++){
				 //sVB = sVB + (j == 0 ? "" : sBreak) + sVBName + "-" + (j+1) + ":  " + aVB[j];
				 sVB = sVB + (j == 0 ? "" : sBreak) + aVB[j];
				 aArray.push(aVB[j])
			}
			aVB.stream = (j == 1) ? sVB.replace(/.+-[0-9]{1,2}:  (.+)$/ig,"$1") : sVB;
			aVB.array = aArray;
			return aVB;
		}
		return false;
	}
	catch(e){
		return false;
	}
}




////////////////////////////////////////////////////////////////////////////////
///////// ERROR FUNCTIONS


function wmi_log_last(){
	try{
		// Create the last error object
		sErr = "\nWMI Last Error Information:";
		sErr += "\nOperation: " + oLastErr.Operation;
		sErr += "\nProvider: " + oLastErr.ProviderName;
	
		if(oLastErr.Description ==  "") sErr += "\nDescription:" + oLastErr.Description;
		if(oLastErr.ParameterInfo == "") sErr += "\nParameter Info: " + oLastErr.ParameterInfo;
		if(oLastErr.StatusCode == "") sErr += "\nStatus: " + oLastErr.StatusCode;
		
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
	finally{
		return sErr;
	}
}

function wmi_log_clear(){
	try{		
		wmi.err.error = null
		wmi.err.number = 0;
		wmi.err.description = "";
	}
	catch(e){
		
	}
}

function wmi_log_error(iOpt,oErr){
	try{		
		js_log_error(iOpt,oErr);
	}
	catch(e){
		try{
			WScript.Echo(oErr.description);
		}
		catch(e){
			alert(oErr.description);
		}
	}
}
