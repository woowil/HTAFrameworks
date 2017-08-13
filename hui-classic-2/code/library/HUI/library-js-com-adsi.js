// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

// http://www.microsoft.com/technet/prodtechnol/windowsserver2003/library/TechRef/54094485-71f6-4be8-8ebf-faa45bc5db4c.mspx
/*

File:     library-js-com-adsi.js
Purpose:  Development script for ADSI/IIS
Author:   Woody Wilson
Created:  2003-10
Version:  see LIB_VERSION

Description:
JScript library scripting functions used for WSH files or HTA Applications.
 
Usage:
Need library-js.js, library-js-wmi.js, library-js-htm.js, library-js-adsi.js, library-js-xml.js, library-js-ado.js

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

// GLOBAL/EXTERNAL DECLARATIONS
try{
	if(!oFso || !oWsh || !oWno){
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oWsh = new ActiveXObject("WScript.Shell");
		var oWno = new ActiveXObject("WScript.Network");
	}
}
catch(e){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
}

var LIB_NAME    = "Library SYS ADSI";
var LIB_VERSION = "1.0";
var LIB_FILE    = oFso.GetAbsolutePathName("library-js-sys-adsi.js")

var adsi = new adsi_object()

function adsi_object(){
	try{
		this.version = new js_objectversion(LIB_NAME,LIB_FILE,LIB_VERSION)
	}
	catch(ee){}
	
}

// Function _Class() must be loaded
class_mgnt_adsi.prototype = new _Class("NT/ADSI","Class for managing Active Directory Users and Computers")

function class_mgnt_adsi(oService,sComputer,sUser,sPass,bInit,bStatus,bDebug){
	try{
		this.wmi_cimv2 = oService
		this.regprov = null
		this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
		this.username = typeof(sUser) == "string" ? sUser : null
		this.password = typeof(sPass) == "string" ? sPass : null
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false				
		this.name = "class_mgnt_adsi"
		this.bInitService = false
		this.oADUser = null
		this.oNTUser = null
		this.sLDAPRoot1 = null
		this.sDNC = null
		var oConnection = null
		var oCommand = null
		var oRecordSet = null
		
		this.getDNC = function(){
			return (GetObject("LDAP://rootDSE")).Get("defaultNamingContext")
		}
		
		this.getConnection = function(){
			try{
				if(!oConnection){
					var ADS_SCOPE_SUBTREE = 2
					// Prevents ADO message like: Safety settings on this computer prohibit accessing a data source on another domain.
					// => IE Security->Custom Level->Internet|Local Intranet: Access Data Sources Across Domains
					// http://www.jsifaq.com/SF/Tips/Tip.aspx?id=5130
					// This must be set, or it generates error 2716
					//oWsh.RegWrite(oReg.ie_cu + "\\Zones\\1\\1406",0,"REG_DWORD") // Intranet zone
					//oWsh.RegWrite(oReg.ie_cu + "\\Zones\\3\\1406",0,"REG_DWORD") // Internet zone
					oConnection = new ActiveXObject("ADODB.Connection")					
					oCommand = new ActiveXObject("ADODB.Command")
					oConnection.Provider = "ADsDSOObject"
					oConnection.Open("Active Directory Provider")
					oCommand.ActiveConnection = oConnection
					oCommand.Properties("Page Size") = 1000
					oCommand.Properties("Timeout") = 30
					oCommand.Properties("Cache Results") = false
					oCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
					this.sDNC = (GetObject("LDAP://rootDSE")).Get("defaultNamingContext")
				}
				return true
			}
			catch(ee){
				// 2716: Safety settings on this computer prohibit accessing a data source on another domain.
				// 3709: The connection cannot be used to perform this operation. It is either closed or invalid in this context.
				this.error(ee,"getConnection()");
				return false
			}
		}
		
		this.getUser = function(sUser,sLDAP,sFilter){
			try{
				if(!this.getConnection()) return false
				sLDAP = typeof(sLDAP) == "string" ? sLDAP : "LDAP://" + this.sDNC
				/*
				var sBase = "<" + sLDAP + ">"
				var sFilter = sFilter ? sFilter : "(&(objectCategory=person)(objectClass=user)(samAccountName=" + sUser + "))"
				var sAttributes = "distinguishedName,samAccountName"
				var sQuery = sBase + ";" + sFilter + ";" + sAttributes + ";subtree"
				*/
				oCommand.CommandText = "SELECT distinguishedName FROM '" + sLDAP + "' WHERE objectCategory='user' AND samAccountName = '" + sUser + "'"
				oRecordSet = oCommand.Execute()
				//alert(oRecordSet.RecordCount)
				// Do not use oRecordSet.MoveFirst(), generates error 3021
				for(; !oRecordSet.EOF; oRecordSet.MoveNext()){
					return GetObject("LDAP://" + oRecordSet.Fields("distinguishedName"))
				}
			}
			catch(ee){
				// 3021: Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.
				//       http://www.scit.wlv.ac.uk/appdocs/chili-asp/html/ado_recordset_object_bof_eof_properties.htm
				this.error(ee,"getUser()");				
			}
			return null
		}
		
		/*
		Mandatory attributes:
1       cn = SM82558
2       instanceType = 4
4       objectCategory = CN=Computer,CN=Schema,CN=Configuration,DC=sdproot,DC=net
7       sAMAccountName = SM82558$

Optional attributes:
30      badPwdCount = 0
39      codePage = 0
43      countryCode = 0
61      distinguishedName = CN=SM82558,OU=PC15,OU=Clients,OU=A1,DC=dnbnor,DC=net
69      dNSHostName = sm82558.dnbnor.net
141     isCriticalSystemObject = False
157     localPolicyFlags = 0
160     logonCount = 34
285     name = SM82558
298     operatingSystem = Windows XP Professional
300     operatingSystemServicePack = Service Pack 2
301     operatingSystemVersion = 5.1 (2600)
332     primaryGroupID = 515
356     sAMAccountType = 805306369
404     userAccountControl = 4096
425     whenChanged = 07.04.2009 15:15:11
426     whenCreated = 05.06.2008 08:53:37
		*/

		this.getComputer = function(sComputer,sLDAP,sFilter){
			try{
				if(!this.getConnection()) return false
				sLDAP = typeof(sLDAP) == "string" ? sLDAP : "LDAP://" + this.sDNC
				/*
				var sBase = "<" + sLDAP + ">"
				var sFilter = sFilter ? sFilter : "(&(objectCategory=person)(objectClass=user)(samAccountName=" + sUser + "))"
				var sAttributes = "distinguishedName,samAccountName,cn,dnsHostName"
				var sQuery = sBase + ";" + sFilter + ";" + sAttributes + ";subtree"
				*/
				oCommand.CommandText = "SELECT distinguishedName FROM '" + sLDAP + "' WHERE objectCategory='computer' AND name='" + sComputer + "'"
				oRecordSet = oCommand.Execute()
				//alert(oCommand.CommandText)
				//alert(oRecordSet.RecordCount)
				// Do not use oRecordSet.MoveFirst(), generates error 3021
				for(; !oRecordSet.EOF; oRecordSet.MoveNext()){
					return GetObject("LDAP://" + oRecordSet.Fields("distinguishedName"))
				}
			}
			catch(ee){
				// 3021: Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.
				//       http://www.scit.wlv.ac.uk/appdocs/chili-asp/html/ado_recordset_object_bof_eof_properties.htm
				this.error(ee,"getComputer()");
			}
			return null
		}
		
		this.getComputerDN = function(sComputer){
			var oo
			if((oo = this.getComputer(sComputer))){
				return oo.distinguishedName
			}
			return undefined
		}
		
		this.getUserAccount = function(sUser,sLDAP){
			try{
				this.echo("# Getting user account for " + sUser);
				this.reset()
				this.oADUser = this.getUserLDAP(sUser,sLDAP)
				return this.oADUser
			}
			catch(ee){
				this.error(ee,"getUserAccount()");				
			}
			return null
		}
		
		////////////// SET
		/////////////////////////////////////////////////////////////////////////////////
		
		this.setUserDisabled = function(sUser,oUser){
			return;
			if(typeof(oUser) == "object");
			else if(oUser = this.getUserAccount(sUser));
			else return false
			if(this.oADUser){
				var ADS_UF_ACCOUNTDISABLE = 2
				var iUAC = this.oADUser.Get("userAccountControl"); 
				this.oADUser.Put("userAccountControl",iUAC || ADS_UF_ACCOUNTDISABLE);
				this.oADUser.SetInfo();
				return true
			}
			return false
		}
		
		this.setUserEnabled = function(sUser){
			return;
			if(typeof(oUser) == "object");
			else if(oUser = this.getSaturnUser(sUser));
			else return false
			if(this.oNTUser){
				this.oNTUser.AccountDisabled = false;
				this.oNTUser.SetInfo();
				return true
			}
			return false
		}
		
		this.setUserUnlock = function(sUser,oUser){
			return;
			if(typeof(oUser) == "object");
			else if(oUser = this.getSaturnUser(sUser));
			else return false
			if(this.oNTUser){
				this.oNTUser.IsAccountLocked = false;
				this.oNTUser.SetInfo();
				return true
			}
			return false
		}
		
		this.setUserChangePwd = function(sUser,oUser){
			return;
			if(typeof(oUser) == "object");
			else if(oUser = this.getSaturnUser(sUser));
			else return false
			if(this.oNTUser){
				this.oNTUser.Put("pwdLastSet",0);
				this.oNTUser.SetInfo();
				return true
			}
			return false
		}
		
		this.reset = function(){
			this.err_number = this.err_description = this.err_function = ""
			this.oADUser = null
			this.oNTUser = null
		}
		
		this.close = function(){
			if(oConnection) oConnection.close(), oConnection = null
			if(oRecordSet) oRecordSet.close(), oRecordSet = null
			//oWsh.RegWrite(oReg.ie_cu + "\\Zones\\1\\1406",1,"REG_DWORD") // Intranet zone
			//oWsh.RegWrite(oReg.ie_cu + "\\Zones\\3\\1406",3,"REG_DWORD") // Internet zone
		}
		
		if(bInit) this.init(oService)
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function adsi_getobject(sGetObject,bError){
	try{
		return GetObject(sGetObject);
	}
	catch(e){
		if(bError) adsi_log_error(2,e);
		return false;
	}
}

function adsi_getinfo(sGetObject,bError){
	try{
		var oGetObject = false;
		if(oGetObject = GetObject(sGetObject)){
			try{
				oGetObject.GetInfo(); 
			}
			catch(ee){
				// the Active Directory property cannot be found in the cache
				// http://www.windowsitpro.com/Articles/Index.cfm?ArticleID=15734&pg=2
				oGetObject.SetInfo()
				oGetObject.GetInfoEx();
			}
		}
		return oGetObject;
	}
	catch(e){
		if(bError) adsi_log_error(2,e);
		return false;
	}
}

function adsi_getuser(sUser,sLDAP,sFilter){
	try{
		//if(!oConnection || !oCommand){
			var ADS_SCOPE_SUBTREE = 2
			oConnection = new ActiveXObject("ADODB.Connection")
			oCommand = new ActiveXObject("ADODB.Command")
			oConnection.Provider = "ADsDSOObject"
			oConnection.Open("Active Directory Provider")
			oCommand.ActiveConnection = oConnection

			oCommand.Properties("Page Size") = 100
			oCommand.Properties("Timeout") = 20
			oCommand.Properties("Cache Results") = false
			oCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
			oDNC = (GetObject("LDAP://rootDSE")).Get("defaultNamingContext")
		//}
		sLDAP = typeof(sLDAP) == "string" ? sLDAP : "LDAP://" + oDNC
		
		var sBase = "<" + sLDAP + ">"
		var sFilter = sFilter ? sFilter : "(&(objectCategory=person)(objectClass=user)(samAccountName=" + sUser + "))"
		var sAttributes = "distinguishedName,samAccountName"
		var sQuery = sBase + ";" + sFilter + ";" + sAttributes + ";subtree"
				
		oCommand.CommandText = "SELECT distinguishedName FROM '" + sLDAP + "' WHERE objectCategory='user' AND samAccountName = '" + sUser + "'"
		var oRecordSet = oCommand.Execute()		
		
		for(; !oRecordSet.EOF; oRecordSet.MoveNext()){
			return GetObject("LDAP://" + oRecordSet.Fields("distinguishedName").Value)
		}
		return null
	}
	catch(e){
		error(e,true)
		return false;
	}
	finally{
		try{
			oRecordSet.close()
			oConnection.close()
		}
		catch(ee){}
	}
}

function adsi_isingroup(sOpt,sGetObject,sMember){
	try{
		var bInGroup = false;
		if(sOpt == "nt"){ // Windows NT Domain group or local groups
			var oRe = /WinNT:\/\/([a-z. \-]+)\/.+,[ \t]*group/ig;
			if(!sGetObject.match(oRe)) return false
			var sWinNT = "WinNT://" + RegExp.$1 + "/" + sMember
			var oGroup = adsi_getobject(sGetObject); // "WinNT://DNB/7074,group"
			if(oGroup){
				return oGroup.IsMember(sWinNT);
				/*
				oRe = new RegExp(sMember,"ig");
				for(var oEnum = new Enumerator(oGroup.Members()); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					if((oItem.Name).match(oRe)){
						bInGroup = true;
						break;
					}
				}
				oEnum = null
				*/
			}
		}
		else if(sOpt == "ad"){ // Active Directory groups
			var oGroup = adsi_getinfo(sGetObject,true); // "LDAP://CN="GroupName,CN=Users,DC=nor,DC=no"			
			var oMemberArray = oGroup.GetEx("member");
			var aMembers = new VBArray(oMemberArray).toArray();
			var oRe = new RegExp(sMember,"ig");
			for(var i = 0, len = aMembers.length; i < len; i++){
				if(aMembers[i].match(oRe)){
					bInGroup = true;
					break;
				}
			}
		}
		else {
			return sat_isingroup("nt",sGetObject,sMember) || sat_isingroup("ad",sGetObject,sMember);
		}
		return bInGroup;
	}
	catch(e){
		sat_log_error(2,e);
		return false;
	}
}

function adsi_dse_user(sOpt,sGetObject,oUser){
	try{
		if(sOpt == "getgroups"){ // Active Directory groups
			oUser = js_str_isdefined(oUser) ? oUser : adsi_getinfo(sGetObject,true); // "LDAP://CN="GroupName,CN=Users,DC=nor,DC=no"
			var aGroups = new Array()
			try{
				var oMembers = oUser.Groups()
				for(var oItem, oEnum = new Enumerator(oMembers); !oEnum.atEnd(); oEnum.moveNext()){
					oItem = oEnum.item();
					var o = new Object()
					o.Name = (new String(oItem.Name)).replace(/cn=(.+)/ig,"$1")
					o.DistinguishedName = oItem.distinguishedName
					aGroups.push(o)
				}
				htm_ary_sort(2,aGroups,"Name")
			}
			catch(ee){
				var oMembers = oUser.Members()
				aGroups = (new VBArray(oMembers).toArray());
				aGroups.sort()
			}
			return aGroups
		}
		return null;
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
}

function adsi_dse_getfsmo(){
	try{
		this.root = adsi_getobject("LDAP://rootDSE")
		if(!this.root) return null
		this.snc = this.root.Get("schemaNamingContext") // 
		this.dnc = this.root.Get("defaultNamingContext") // 
		this.cnc = this.root.Get("configurationNamingContext")
		// Forest-wide Schema Master FSMO
		var oSchema = adsi_getobject("LDAP://" + this.snc)
		sSchema = oSchema.Get("fSMORoleOwner")
		var oNtds = adsi_getobject("LDAP://" + sSchema)
		var oComputer = adsi_getobject(oNtds.Parent)
		this.schemamaster =  (oComputer.Name).replace(/cn=/ig,"")
		// Forest-wide Domain Naming Master FSMO
		var oPartition = adsi_getobject("LDAP://CN=Partitions," + this.cnc)
		sNaming = oPartition.Get("fSMORoleOwner")
		oNtds = adsi_getobject("LDAP://" + sNaming)
		oComputer = adsi_getobject(oNtds.Parent)
		this.namingmaster = (oComputer.Name).replace(/cn=/ig,"")
		// Domain's PDC Emulator FSMO
		var oDomain = adsi_getobject("LDAP://" + this.dnc)
		sDomain = oDomain.Get("fSMORoleOwner")
		oNtds = adsi_getobject("LDAP://" + sDomain)
		oComputer = adsi_getobject(oNtds.Parent)
		this.pdcemulator = (oComputer.Name).replace(/cn=/ig,"")
		// Domain's RID Master FSMO
		var oDomain = adsi_getobject("LDAP://CN=RID Manager$,CN=System," + this.dnc)
		sDomain = oDomain.Get("fSMORoleOwner")
		oNtds = adsi_getobject("LDAP://" + sDomain)
		oComputer = adsi_getobject(oNtds.Parent)
		this.ridmaster = (oComputer.Name).replace(/cn=/ig,"")
		// Domain's Infrastructure Master FSMO
		var oDomain = adsi_getobject("LDAP://CN=Infrastructure," + this.dnc)
		sDomain = oDomain.Get("fSMORoleOwner")
		oNtds = adsi_getobject("LDAP://" + sDomain)
		oComputer = adsi_getobject(oNtds.Parent)
		this.infrastructure = (oComputer.Name).replace(/cn=/ig,"")
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
}

/*


Set oUClass = GetObject("LDAP://schema/printQueue")
'Set oUClass = GetObject("LDAP://schema/user")
'Set oUClass = GetObject("LDAP://schema/computer")
'Set oUClass = GetObject("LDAP://schema/contacts")
Set oSchema = GetObject(oUClass.Parent)

WScript.echo "","[Mandatory atrributes]"
For Each GroupObj In oUClass.MandatoryProperties
	Wscript.Echo GroupObj

Next

WScript.echo "","[Optional attributes]"
For Each GroupObj In oUClass.OptionalProperties
	Wscript.Echo GroupObj

Next
*/

function adsi_adodb_object(sProvider,sProviderOpen,sCommandText,oProperties){
		// http://www.scit.wlv.ac.uk/appdocs/chili-asp/html/ado_recordset_object_bof_eof_properties.htm
		this.connect = new ActiveXObject("ADODB.Connection");
		this.command = new ActiveXObject("ADODB.Command")
		this.provider = this.connect.Provider = sProvider ? sProvider : "ADsDSOObject";
		this.connect.Open(sProviderOpen); // Active Directory Provider
		this.command.ActiveConnection = this.connect;
		this.commandtext = this.command.CommandText = sCommandText;
		if(typeof(oProperties) == "object"){
			for(sProperty in oProperties){
				this.command.Properties(sProperty) = oProperties[sProperty];
			}
		}
		//this.execute = this.command.Execute;
		this.recordset = null; // this.command.Execute();
}

function adsi_adodb_ldap(sOpt1,sOpt2,sOpt3,sLDAP){
	try{
		var oAdo;
		var sProvider = "ADsDSOObject"
		var sProviderOpen = "Active Directory Provider"
		var sCommandText = ""
		//if(!js_str_isdefined(sOpt1,sLDAP)) return false;
		if(sOpt1 == "computers"){
			var ADS_SCOPE_SUBTREE = !isNaN(arguments[2]) ? arguments[2] : 2; // 2 = Recursivt OU
			var MAX_COMPUTERS = !isNaN(arguments[3]) ? arguments[3] : 2000;
			sCommandText = "Select Name, Location from 'LDAP://" + sLDAP + "' Where objectClass='computer'" // sLDAP: OU=KS10,OU=Servers,OU=A1,DC=dnbnor,DC=net
			var oComputers = new Array();
			var oProperties = new Array();
			oProperties["Page Size"] = MAX_COMPUTERS
			oProperties["Searchscope"] = ADS_SCOPE_SUBTREE
			oAdo = new adsi_adodb_object(sProvider,sProviderOpen,sCommandText,oProperties);
			oAdo.recordset = oAdo.command.Execute();
			oAdo.recordset.MoveFirst();
			for(var i = 0; !oAdo.recordset.EOF; i++, oAdo.recordset.MoveNext()){
				oComputers[i] = new Object()
				oComputers[i].Name = oAdo.recordset.Fields("Name").Value;
				oComputers[i].Location = oAdo.recordset.Fields("Location").Value;
			}
			return oComputers;
		}
		else if(sOpt1 == "getuser"){
			
		}
		else if(sOpt1 == "printer"){
			/* dnbnor.net/A1/Servers/KS10/CFABK001N/CFABK001N-KIG-048
			curComputerADSite=GAW
			curComputerADSiteDescr=nedreskoyenvei
			curComputerADSubnet=10.226.11.0/24
			curComputerADSubnetLocation=GAW
			CurComputerDHCPDomain=Lorenveien.DnBnor.lan
			curComputerDNSDomain=dnbnor.net
			*/
			sDomain = sLDAP
			sADSubnetLocation = sOpt3			
			if(sOpt2 == "bc10") sADSubnetLocation += "/*"
			else if(sOpt2 == "gc10");
			else return false
			var oPrinters = new Array();
			sCommandText = "Select printerName, serverName, Location, Comment from 'LDAP://" + sDomain + "' where objectClass='printQueue' and Location = '" + sADSubnetLocation + "'"
			oAdo = new adsi_adodb_object(sProvider,sProviderOpen,sCommandText);
			oAdo.recordset = oAdo.command.Execute();
			oAdo.recordset.MoveFirst();
			for(var i = 0; !oAdo.recordset.EOF; i++, oAdo.recordset.MoveNext()){
				oPrinters[i] = new Object()
				oPrinters[i].Server = oAdo.recordset.Fields("serverName").Value; // clstk001n.dnbnor.net
				oPrinters[i].Printer = oAdo.recordset.Fields("printerName").Value; // LST-001
				oPrinters[i].Comment = oAdo.recordset.Fields("Comment").Value;
				oPrinters[i].PrinterShare = "\\\\" + oPrinters[i].Server + "\\" + oPrinters[i].Printer; // \\clstk001n.dnbnor.net\LST-001
			}
			return oPrinters;
		}
		else if(sOpt1 == "printer_loren"){
			sCommandText = "Select printerName, serverName, Location from 'LDAP://cgawd010n.dnbnor.net/OU=Servers,OU=A1,DC=dnbnor,DC=net' where objectClass='printQueue'"
			oAdo = new adsi_adodb_object(sProvider,sProviderOpen,sCommandText);
			oAdo.recordset = oAdo.command.Execute();
			oAdo.recordset.MoveFirst();
			
			for(var o; !oAdo.recordset.EOF; oAdo.recordset.MoveNext()){
				var p = new String(oAdo.recordset.Fields("printerName").Value)
				var ll = new String(oAdo.recordset.Fields("Location").Value)
				var s = (new String(oAdo.recordset.Fields("serverName").Value)).replace(/([\w\-])\..+/,"$1")
				var l = "LDAP://cgawd010n.dnbnor.net/CN=" + s + "-" + p + ",CN=" + s + ",OU=KS10,OU=Servers,OU=A1,DC=dnbnor,DC=net"
				if(ll.match(/gaw/ig) && (o = adsi_getobject(l))){
					if(o.Description && (o.Description).match(/faret|nveien/ig)){
						 ss = o.Location + ";"
						 ss = ss + s + " ; "
						 ss = ss + p + " ; "; // LST-001
						 ss = ss + o.Description;
						 prn(ss)
					 }
				}
				//
			}
		}
		return false
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
	finally{
		adsi_adodb_close(oAdo);
	}
}

function adsi_adodb_close(oADODB){
	try{
		try{
			oADODB.connect.Close();
			oADODB.recordset = null
		}
		catch(ee){}
		delete oADODB;
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
}

function sat_adsi_addgroup(sComputer,sLocalGroup,oGroup){
	try{
		sComputer = sComputer ? sComputer : oWno.ComputerName;
		sLocalGroup = sLocalGroup ? sLocalGroup : "Administrators";
		var objGroup = GetObject("WinNT://" + sComputer + "/" + sLocalGroup + ",group");
		objGroup.Add(oGroup.ADsPath);
		return true;
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
}

/////// IIS ADSI
// IIsWebServer: http://feurerpc.oet.udel.edu/iishelp/iis/htm/asp/aore7zn6.htm
// IIsFilters: http://feurerpc.oet.udel.edu/iishelp/iis/htm/asp/aore4h6b.htm
// IIsCertMapper: http://feurerpc.oet.udel.edu/iishelp/iis/htm/asp/aore4qia.htm

function adsi_iis_ok(sDomain){
	try{
		sDomain = sDomain ? sDomain : oWno.ComputerName;
		var oIisService = GetObject("IIS://" + sDomain + "/W3SVC");
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
	finally{
		return oIisService;
	}
}

function adsi_iis_vircreate(sDomain,sWebSite,sVirPath,sVirName,oIisService){
	try{
		oIisService = oIisService ? oIisService : adsi_iis_ok(sDomain);
		if(typeof(oIisService) == "object"){
			// Check if the Web site directory exists
			for(var oItem, bWebID = false, oEnum = new Enumerator(oIisService); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.Class == "IIsWebServer"){
					if(oItem.ServerComment == sWebSite){
						bWebID = oItem.Name; // 1..n
						break;
					}
				}
			}
			if(!bWebID){
				var sMsg = "TERMINATING CreateWebVirtualDir, Error accessing Web Site '" + sWebSite + "'. Server may not exist.";
				js.err.description = sMsg;
				return false;
			}
			
			// Try to open the web site ID.
			try{
				oWebSite = oIisService.GetObject("IIsWebServer",bWebID);
				// Get the web site's virtual root
				var oVRoot = oWebSite.GetObject("IIsWebVirtualDir","Root");
				try{
					// Create the new virtual directory if not exist
					var oVDir = oVRoot.Create("IIsWebVirtualDir",sVirName);
				}
				catch(eee){}
				// Sets and saves the new virtual directory path
				oVDir.AccessRead = true;
				oVDir.Path = sVirPath;
				oVDir.SetInfo();
			}
			catch(ee){
				adsi_log_error(2,ee);
				return false;
			}
			
			return true;
		}
		else return false;
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
}

function adsi_iis_virdelete(sDomain,sWebSite,sVirName,oIisService){
	try{
		oIisService = oIisService ? oIisService : adsi_iis_ok(sDomain);
		if(typeof(oIisService) == "object"){
			// Check if the Web site and virtual exists
			for(var oItem, bWebID = false, oEnum = new Enumerator(oIisService); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.Class == "IIsWebServer"){
					if(oItem.ServerComment == sWebSite){						
						bWebID = oItem.Name; // 1..n
						break;
					}
				}
			}
			if(!bWebID){
				adsi_log_error(2,ee);
				return false;
			}
			
			// Try to open the web site ID.
			try{
				oWebSite = oIisService.GetObject("IIsWebServer",bWebID);
				// Get the web site's virtual root
				var oVRoot = oWebSite.GetObject("IIsWebVirtualDir","Root");
				// Deletes the virtual directory
				try{
					oVRoot.Delete("IIsWebVirtualDir",sVirName);
				}
				catch(eee){}
			}
			catch(ee){
				adsi_log_error(2,ee);
				return false;
			}
			return true;
		}
		else return false;
	}
	catch(e){
		adsi_log_error(2,e);
		return false;
	}
}

/////////// ERROR

function adsi_log_error(iOpt,oErr){
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

function adsi_log_error(iOpt,oErr){
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