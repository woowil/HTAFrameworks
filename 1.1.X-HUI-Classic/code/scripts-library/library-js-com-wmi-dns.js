// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-com-wmi-dns.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2005-08
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

///////////////////////
////////// SATURN DNS

// http://msdn.microsoft.com/library/?url=/library/en-us/dns/dns/microsoftdns_cnametype.asp
// http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/networking/dns/

var wmi_dns = new Object()
	wmi_dns.TTL = 600
	wmi_dns.RecordClass = 1
	
function wmi_dns_connect(oService,sServer,sUser,sPass){
	try{
		return oService = oService ? oService : wmi_wbem_service(sServer,"root\\MicrosoftDNS",sUser,sPass);
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_getclass(oService,sClass,stext,sServer,sUser,sPass){
	try{
		oService = oService ? oService : wmi_dns_connect(null,sServer,sUser,sPass)
		return oService.Get(sClass);
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_zone(sOpt,oService,sContainerName,sDomainName){
	try{
		var aClass, sQuery
		var sClass = "MicrosoftDNS_Zone"
		sDomainName = sDomainName ? sDomainName : sContainerName
		if(!oService) return false
		if(sOpt == "refresh"){
			sQuery = "Select Name from " + sClass + " where Name='" + sContainerName + "'"
			var oColItems = oService.ExecQuery(sQuery,null,48);
			for(var oItem, i = 0, oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext(), i++){
				oItem = oEnum.item();
				oItem.ForceRefresh()
			}
			return true
		}
		return false;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_domain(sOpt,oService,sContainerName,sDomainName,sRegIgnore){
	try{
		var aClass, sQuery		
		var sClass = "MicrosoftDNS_Domain"
		sDomainName = sDomainName ? sDomainName : sContainerName
		if(!oService) return false
		if(sOpt == "getdomains"){
			//Select Name from MicrosoftDNS_Domain where ContainerName='DnBNOR.LAN'			
			sQuery = "Select Name from " + sClass + " where ContainerName='" + sContainerName + "'" // Ex: ContainerName=Zone=DnBNOR.LAN
			var aSubDomains = new Array(), oRe = new RegExp("([a-z0-9_\-]{2,})\." + sContainerName,"ig"); // skip subdomains with numeric
			var oRe2 = new RegExp(new String(sRegIgnore),"ig"), o, oo
			if(aClass = wmi_win32_class("simple",sClass,oService,sQuery,"\n")){
				for(var i = j = 0, len = aClass.length; i < len && !js.pro.stopit; i++, j = 0){
					for(o in aClass[i]){
						oo = aClass[i][o], sSubDomain = new String(oo.value); // honefoss.DnBNOR.LAN
						if(sSubDomain.match(oRe)){
							var sSubDomain = RegExp.$1
							if(sRegIgnore && sSubDomain.match(oRe2)) continue // skip these subdomains
							aSubDomains.push(js_str_capitalize(sSubDomain));
							if((j++ % 13) == 0) js_tme_sleep(js.time.mil3);
						}
					}					
					aClass[i] = null
				}
			}
			return aSubDomains
		}
		return false;
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_txttype(sOpt,oService,sContainerName,sDomainName,sServer,sText,sOwner){
	try{
		sDomainName = sDomainName ? sDomainName : sContainerName
		var sClass = "MicrosoftDNS_TXTType"
		if(sOpt == "gettxttypes"){			
			sQuery = "Select ContainerName,DescriptiveText,DnsServerName,DomainName,OwnerName,RecordData from " + sClass + " where ContainerName='" + sContainerName + "' And DomainName='" + sDomainName + "'"
			/*
				Caption: null
				ContainerName: DnBNOR.LAN
				Description: null
				DescriptiveText: Site Name: Aas
				DnsServerName: dnbls501.DnB.no
				DomainName: Aas
				InstallDate: undefined
				Name: null
				OwnerName: Aas
				RecordClass: 1
				RecordData: "Site Name: Aas"
				Status: null
				TextRepresentation: Aas IN TXT "Site Name: Aas"
				Timestamp: null
				TTL: 3600
				
				Caption: null
				ContainerName: Printer.DnBNOR.LAN
				Description: null
				DescriptiveText: Model: HP Laserjet 5Si
				DnsServerName: dnbls501.DnB.no
				DomainName: Printer.DnBNOR.LAN
				InstallDate: undefined
				Name: null
				OwnerName: xxx-001.Printer.DnBNOR.LAN
				RecordClass: 1
				RecordData: "Model: HP Laserjet 5Si"
				Status: null
				TextRepresentation: xxx-001.Printer.DnBNOR.LAN IN TXT "Model: HP Laserjet 5Si"
				Timestamp: null
				TTL: 3600

			*/
			return wmi_win32_class("simple",sClass,oService,sQuery,"\n");
		}
		else if(sOpt == "gettxttypesall"){			
			if(aDomains = wmi_dns_domain("getdomains",oService,sContainerName,sDomainName)){
				var aOClass = new Array()
				for(var i = 0, len = aDomains.length; i < len; i++){
					if(aClass = wmi_dns_domain("gettxttypes",oService,sContainerName,aDomains[i])){
						aOClass.push(aClass)
						aClass = null
					}
				}
				return aOClass
			}
		}
		else if(sOpt == "deltxttype"){
			var oClass = wmi_dns_getclass(oService,sClass)
			sOwner = sOwner ? sOwner : sServer + "." + sContainerName
			var oResult = oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,wmi_dns.RecordClass,wmi_dns.TTL,sText)			
			
		}
		else if(sOpt == "addtxttype"){
			var oClass = wmi_dns_getclass(oService,sClass)
			sOwner = sOwner ? sOwner : sServer + "." + sContainerName
			sText = js_str_trim(sText)
			var sText2 = sText.replace(/\\/g,"\\\\") // Bug!!
			try{
				oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,wmi_dns.RecordClass,wmi_dns.TTL,sText2);
			}
			catch(ee){} // Generic failure
			return (oWsh.Run("%comspec% /c nslookup -type=TXT " + sOwner + " " + sServer + " | find /i \"Non-existent\" >nul",oReg.hide,true) != 0);	
		}
		return false
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_atype(sOpt,oService,sContainerName,sDomainName,sServer,sIPaddress,sOwner){
	try{
		sDomainName = sDomainName ? sDomainName : sContainerName
		var sClass = "MicrosoftDNS_AType"
		if(sOpt == "getatypes"){			
			sQuery = "Select ContainerName,DescriptiveText,DnsServerName,DomainName,OwnerName,RecordData from " + sClass + " where ContainerName='" + sContainerName + "' And DomainName='" + sDomainName + "'"
			/*
				
			*/
			return wmi_win32_class("simple",sClass,oService,sQuery,"\n");
		}
		else if(sOpt == "getatypesall"){
			if(aDomains = wmi_dns_domain("getdomains",oService,sContainerName,sDomainName)){
				var aOClass = new Array()
				for(var i = 0, len = aDomains.length; i < len; i++){
					if(aClass = wmi_dns_domain("gettxttypes",oService,sContainerName,aDomains[i])){
						aOClass.push(aClass)
						aClass = null
					}
				}
				return aOClass
			}
		}
		else if(sOpt == "delatype"){
			var oClass = wmi_dns_getclass(oService,sClass)
			sOwner = sOwner ? sOwner : sServer + "." + sContainerName
			var oResult = oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,wmi_dns.RecordClass,wmi_dns.TTL,sText)
		}
		else if(sOpt == "addatype"){
			var oClass = wmi_dns_getclass(oService,sClass)
			sOwner = sOwner ? sOwner : sServer + "." + sContainerName
			try{
				oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,wmi_dns.RecordClass,wmi_dns.TTL,sIPaddress);
			}
			catch(ee){} // Generic failure
			return (oWsh.Run("%comspec% /c nslookup -type=A " + sOwner + " " + sServer + " | find /i \"Non-existent\" >nul",oReg.hide,true) != 0);
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_cnametype(sOpt,oService,sContainerName,sDomainName,sServer,sOwner,sPrimaryName){
	try{
		sDomainName = sDomainName ? sDomainName : sContainerName
		var oClass
		var sClass = "MicrosoftDNS_CNAMEType"
		if(sOpt == "getcnametypes"){
			sQuery = "Select ContainerName,DnsServerName,DomainName,OwnerName,PrimaryName,RecordClass,RecordData,TextRepresentation from " + sClass + " where ContainerName='" + sContainerName + "' And DomainName='" + sDomainName + "'"
			/*
				Caption: null
				ContainerName: DnBNOR.LAN
				Description: null
				DnsServerName: dnbls501.DnB.no
				DomainName: Aas
				InstallDate: undefined
				Name: null
				OwnerName: local-distribution-server.Aas
				PrimaryName: cahgk001n.dnbnor.net.
				RecordClass: 1
				RecordData: cahgk001n.dnbnor.net.
				Status: null
				TextRepresentation: local-distribution-server.Aas IN CNAME cahgk001n.dnbnor.net.
				Timestamp: null
				TTL: 3600
			*/
			return wmi_win32_class("simple",sClass,oService,sQuery,"\n");
		}
		else if(sOpt == "getcnametypesall"){			
			if(aDomains = wmi_dns_domain("getdomains",oService,sContainerName,sDomainName)){
				var aOClass = new Array()
				for(var i = 0, len = aDomains.length; i < len; i++){
					if(aClass = wmi_dns_domain("getcnametypes",oService,sContainerName,aDomains[i])){
						aOClass.push(aClass)
						aClass = null
					}
				}
				return aOClass
			}
		}		
		else if(sOpt == "delcname"){
			if(oClass = wmi_dns_getclass(oService,sClass)){
				sOwner = sOwner ? sOwner : sServer + "." + sContainerName
				var oResult = oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,wmi_dns.RecordClass,wmi_dns.TTL,sText)
			}			
		}
		else if(sOpt == "addcname"){
			if(oClass = wmi_dns_getclass(oService,sClass)){
				var oRe = new RegExp("[a-z0-9\-_]+\." + sDomainName + "$","ig"); // Must have ".Domain" 
				sOwner = sOwner.match(oRe) ? sOwner : sOwner + "." + sDomainName // local-printer-server.Domain
				try{
					oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,wmi_dns.RecordClass,wmi_dns.TTL,sPrimaryName);
				}
				catch(ee){} // Generic failure
				return (oWsh.Run("%comspec% /c nslookup -type=CNAME " + sOwner + " " + sServer + " | find /i \"Non-existent\" >nul",oReg.hide,true) != 0);
			}
		}
		else if(sOpt == "setcname"){
			/*
			if(oClass = wmi_dns_getclass(oService,sClass)){
				sOwner = sOwner ? sOwner : sServer + "." + sContainerName
				oClass.Modify(wmi_dns.TTL,sPrimaryName)
			}
			*/
		}
		return false
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

function wmi_dns_soatype(sOpt,oService,sContainerName,sDomainName,sServer,sOwner,sPrimaryName){
	try{
		sDomainName = sDomainName ? sDomainName : sContainerName
		var oClass, sQuery
		var sClass = "MicrosoftDNS_SOAType"
		if(sOpt == "getserial"){
			sQuery = "Select SerialNumber from " + sClass + " where ContainerName='" + sContainerName + "'"
			if(oClass = wmi_win32_class("simple",sClass,oService,sQuery,"\n")){
				for(o in oClass[0]){
					return oClass[0][o].value
				}
			}
		}
		return false
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}