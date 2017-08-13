// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


// http://www.microsoft.com/technet/prodtechnol/windowsserver2003/library/TechRef/54094485-71f6-4be8-8ebf-faa45bc5db4c.mspx

__H.include("HUI@Sys@LDAP.js")

__H.register(__H.Sys.LDAP,"ADSI","Directory Services (ADSI)",function ADSI(){
	this.oADUser = null
	this.oNTUser = null
	this.sLDAPRoot1 = null
	var oConnection = null
	var oCommand = null
	var oRecordSet = null
		
	this.getDNC = function getDNC(){
		return (GetObject("LDAP://rootDSE")).Get("defaultNamingContext")
	}
	
	this.getConnection = function getConnection(){
		try{
			if(!oConnection){
				var ADS_SCOPE_SUBTREE = 2
				// Prevents ADO message like: Safety settings on this computer prohibit accessing a data source on another domain.
				// => IE Security->Custom Level->Internet|Local Intranet: Access Data Sources Across Domains
				// http://www.jsifaq.com/SF/Tips/Tip.aspx?id=5130
				// This must be set, or it generates error 2716
				//oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\1\\1406",0,"REG_DWORD") // Intranet zone
				//oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\3\\1406",0,"REG_DWORD") // Internet zone
				oConnection = new ActiveXObject("ADODB.Connection")					
				oCommand = new ActiveXObject("ADODB.Command")
				oConnection.Provider = "ADsDSOObject"
				oConnection.Open("Active Directory Provider")
				oCommand.ActiveConnection = oConnection
				oCommand.Properties("Page Size") = 1000
				oCommand.Properties("Timeout") = 30
				oCommand.Properties("Cache Results") = true
				oCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
				this.sDNC = (GetObject("LDAP://rootDSE")).Get("defaultNamingContext")
			}
			return true
		}
		catch(ee){
			// 2716: Safety settings on this computer prohibit accessing a data source on another domain.
			// 3709: The connection cannot be used to perform this operation. It is either closed or invalid in this context.
			__HLog.error(ee,this);
			return false
		}
	}
	
	this.getUser = function getUser(sUser,sLDAP,sFilter){
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
			__HLog.error(ee,this);				
		}
		return null
	}
	
	this.getComputer = function getComputer(sComputer,sLDAP,sFilter){
		try{
			if(!this.getConnection()) return false
			
			sLDAP = typeof(sLDAP) == "string" ? sLDAP : "LDAP://" + this.sDNC
			/*
			var sBase = "<" + sLDAP + ">"
			var sFilter = sFilter ? sFilter : "(&(objectCategory=person)(objectClass=objectClass)(samAccountName=" + sUser + "))"
			var sAttributes = "distinguishedName,samAccountName"
			var sQuery = sBase + ";" + sFilter + ";" + sAttributes + ";subtree"
			*/
			oCommand.CommandText = "SELECT cn,dnsHostName FROM '" + sLDAP + "' WHERE objectClass='computer' AND sAMAccountName = '" + sComputer + "'"
			oRecordSet = oCommand.Execute()
			//alert(oRecordSet.RecordCount)
			// Do not use oRecordSet.MoveFirst(), generates error 3021
			for(; !oRecordSet.EOF; oRecordSet.MoveNext()){
				//alert(oRecordSet.Fields("cn"))
				//alert(oRecordSet.Fields("dnsHostName"))
				return GetObject("LDAP://" + oRecordSet.Fields("cn"))
			}
		}
		catch(ee){
			// 3021: Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.
			//       http://www.scit.wlv.ac.uk/appdocs/chili-asp/html/ado_recordset_object_bof_eof_properties.htm
			__HLog.error(ee,this);
		}
		return null
	}
	
	this.getUserAccount = function getUserAccount(sUser,sLDAP){
		try{
			__HLog.log("# Getting user account for " + sUser);
			this.reset()
			this.oADUser = this.getUserLDAP(sUser,sLDAP)
			return this.oADUser
		}
		catch(ee){
			__HLog.error(ee,this);
		}
		return null
	}
	
	////////////// SET
	/////////////////////////////////////////////////////////////////////////////////
	
	this.setUserDisabled = function setUserDisabled(sUser,oUser){
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
	
	this.setUserEnabled = function setUserEnabled(sUser){
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
	
	this.setUserUnlock = function setUserUnlock(sUser,oUser){
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
	
	this.setUserChangePwd = function setUserChangePwd(sUser,oUser){
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
	
	this.reset = function reset(){
		this.err_number = this.err_description = this.err_function = ""
		this.oADUser = null
		this.oNTUser = null
	}
	
	this.close = function close(){
		if(oConnection) oConnection.close(), oConnection = null
		if(oRecordSet) oRecordSet.close(), oRecordSet = null
		//oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\1\\1406",1,"REG_DWORD") // Intranet zone
		//oWsh.RegWrite(__HKeys.ie_cu + "\\Zones\\3\\1406",3,"REG_DWORD") // Internet zone
	}
})

function adsi_getobject(sGetObject,bError){
	try{
		return GetObject(sGetObject);
	}
	catch(ee){
		if(bError) __HLog.error(ee,this)
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
	catch(ee){
		if(bError) __HLog.error(ee,this)
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
	catch(ee){
		__HLog.error(ee,this)
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
					if((oItem.Name).isSearch(oRe)){
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
				if(aMembers[i].isSearch(oRe)){
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
	catch(ee){
		__HLog.error(ee,this)
		return false;
	}
}

function adsi_dse_user(sOpt,sGetObject,oUser){
	try{
		if(sOpt == "getgroups"){ // Active Directory groups
			oUser = !__H.isUndefined(oUser) ? oUser : adsi_getinfo(sGetObject,true); // "LDAP://CN="GroupName,CN=Users,DC=nor,DC=no"
			var aGroups = []
			try{
				var oMembers = oUser.Groups()
				for(var oEnum = new Enumerator(oMembers); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					aGroups.push({
						Name : (new String(oItem.Name)).replace(/cn=(.+)/ig,"$1"),
						DistinguishedName : oItem.distinguishedName
					})
				}
				//htm_ary_sort(2,aGroups,"Name")
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
	catch(ee){
		__HLog.error(ee,this)
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
	catch(ee){
		__HLog.error(ee,this)
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
			for(var sProperty in oProperties){
				if(oProperties.hasOwnProperty(sProperty)){
					this.command.Properties(sProperty) = oProperties[sProperty];
				}
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
		if(sOpt1 == "computers"){
			var ADS_SCOPE_SUBTREE = !isNaN(arguments[2]) ? arguments[2] : 2; // 2 = Recursivt OU
			var MAX_COMPUTERS = !isNaN(arguments[3]) ? arguments[3] : 2000;
			sCommandText = "Select Name, Location from 'LDAP://" + sLDAP + "' Where objectClass='computer'" // sLDAP: OU=KS10,OU=Servers,OU=A1,DC=dnbnor,DC=net
			var oComputers = [];
			var oProperties = [];
			oProperties["Page Size"] = MAX_COMPUTERS
			oProperties["Searchscope"] = ADS_SCOPE_SUBTREE
			oAdo = new adsi_adodb_object(sProvider,sProviderOpen,sCommandText,oProperties);
			oAdo.recordset = oAdo.command.Execute();
			oAdo.recordset.MoveFirst();
			for(var i = 0; !oAdo.recordset.EOF; i++, oAdo.recordset.MoveNext()){
				oComputers[i] = {}
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
			var oPrinters = [];
			sCommandText = "Select printerName, serverName, Location, Comment from 'LDAP://" + sDomain + "' where objectClass='printQueue' and Location = '" + sADSubnetLocation + "'"
			oAdo = new adsi_adodb_object(sProvider,sProviderOpen,sCommandText);
			oAdo.recordset = oAdo.command.Execute();
			oAdo.recordset.MoveFirst();
			for(var i = 0; !oAdo.recordset.EOF; i++, oAdo.recordset.MoveNext()){
				oPrinters[i] = {}
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
				if(ll.isSearch(/gaw/ig) && (o = adsi_getobject(l))){
					if(o.Description && (o.Description).isSearch(/faret|nveien/ig)){
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
	catch(ee){
		__HLog.error(ee,this)
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
	catch(ee){
		__HLog.error(ee,this)
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
	catch(ee){
		__HLog.error(ee,this)
		return false;
	}
}

