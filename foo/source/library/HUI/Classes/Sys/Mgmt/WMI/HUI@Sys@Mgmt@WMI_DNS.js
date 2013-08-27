// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

// http://msdn.microsoft.com/library/?url=/library/en-us/dns/dns/microsoftdns_cnametype.asp
// http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/networking/dns/__H

// TODO
__H.include("HUI@Sys@Mgmt@WMI.js")

__H.register(__H.Sys.Mgmt.WMI,"DNS","Domain Naming System",function DNS(){
	var computer = oWno.ComputerName;
	var username = null
	var password = null
	var namespace = "root\\MicrosoftDNS"
	var o_service = null
	
	this.TTL = 600	
	this.RecordClass = 1
	
	this.initialize = function initialize(sServer,sUser,sPass){
		try{
			o_service = this.setServiceWMI(sComputer,namespace,sUser,sPass)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getclass = function getclass(sClass,sServer,sUser,sPass){
		try{
			if(this.initialize(sServer,sUser,sPass)){
				return o_service.Get(sClass);
			}
			return null
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.zone = function zone(sOpt,sContainerName,sDomainName){
		try{
			var aClass, sClass = "MicrosoftDNS_Zone"
			sDomainName = typeof(sDomainName) == "string" ? sDomainName : sContainerName
			if(!o_service) return false
			if(sOpt == "refresh"){
				var sQuery = "Select Name from " + sClass + " where Name='" + sContainerName + "'"
				var oColItems = o_service.ExecQuery(sQuery,null,48);
				for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !__H.$stopit; oEnum.moveNext()){
					var oItem = oEnum.item();
					oItem.ForceRefresh()
				}
				return true
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getDomains = function getDomains(sContainerName,sDomainName,sRegIgnore){
		try{
			var aClass, sQuery		
			var sClass = "MicrosoftDNS_Domain"
			sDomainName = typeof(sDomainName) == "string" ? sDomainName : sContainerName
			if(!o_service) return false
			//Select Name from MicrosoftDNS_Domain where ContainerName='Nosliw.LAN'			
			sQuery = "Select Name from " + sClass + " where ContainerName='" + sContainerName + "'" // Ex: ContainerName=Zone=Nosliw.LAN
			var aSubDomains = []
			var oRe = new RegExp("([a-z0-9_\-]{2,})\." + sContainerName,"ig"); // skip subdomains with numeric
			var oRe2 = new RegExp(new String(sRegIgnore),"ig"), o, oo
			var a = this.getPropertyNames(sClass)
			var oColItems = oService.ExecQuery(sQuery,"WQL",48)
			for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !__H.$stopit; oEnum.moveNext()){
				var oItem = oEnum.item();				
				for(var i = 0; i < a.length && !__H.$stopit; i++){
					var sSubDomain = new String(oItem[a[i]]) // honefoss.Nosliw.LAN
					if(sSubDomain.match(oRe)){
						var sSubDomain = RegExp.$1
						if(sRegIgnore && sSubDomain.isSearch(oRe2)) continue // skip these subdomains
						aSubDomains.push(sSubDomain.capitalize());
						break;
					}
					delete a[i]
				}
			}
			return aSubDomains
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.txttype = function txttype(sOpt,sContainerName,sDomainName,sServer,sText,sOwner){
		try{
			sDomainName = typeof(sDomainName) == "string" ? sDomainName : sContainerName
			var sClass = "MicrosoftDNS_TXTType"
			if(sOpt == "gettxttypes"){			
				sQuery = "Select ContainerName,DescriptiveText,DnsServerName,DomainName,OwnerName,RecordData from " + sClass + " where ContainerName='" + sContainerName + "' And DomainName='" + sDomainName + "'"
				/*
					Caption: null
					ContainerName: Nosliw.LAN
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
					ContainerName: Printer.Nosliw.LAN
					Description: null
					DescriptiveText: Model: HP Laserjet 5Si
					DnsServerName: dnbls501.DnB.no
					DomainName: Printer.Nosliw.LAN
					InstallDate: undefined
					Name: null
					OwnerName: xxx-001.Printer.Nosliw.LAN
					RecordClass: 1
					RecordData: "Model: HP Laserjet 5Si"
					Status: null
					TextRepresentation: xxx-001.Printer.Nosliw.LAN IN TXT "Model: HP Laserjet 5Si"
					Timestamp: null
					TTL: 3600

				*/
				return wmi_win32_class("simple",sClass,sQuery,"\n",server,namespace,username,password);
			}
			else if(sOpt == "gettxttypesall"){			
				if(aDomains = this.getDomains(sContainerName,sDomainName)){
					var aOClass = []
					for(var i = 0, iLen = aDomains.length; i < iLen; i++){
						if(aClass = this.domain("gettxttypes",sContainerName,aDomains[i])){
							aOClass.push(aClass)						
						}
						__HUtil.kill(aClass)
					}
					__HUtil.kill(aDomains)
					return aOClass
				}
			}
			else if(sOpt == "deltxttype"){
				var oClass = this.getclass(sClass)
				sOwner = typeof(sOwner) == "string" ? sOwner : sServer + "." + sContainerName
				var oResult = oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwnerthis.RecordClass,this.TTL,sText)			
				
			}
			else if(sOpt == "addtxttype"){
				var oClass = this.getclass(sClass)
				sOwner = !__H.isStringEmpty(sOwner) ? sOwner : sServer + "." + sContainerName
				sText = sText.trim()
				var sText2 = sText.replace(/\\/g,"\\\\") // Bug!!
				try{
					oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,this.RecordClass,this.TTL,sText2);
				}
				catch(ee){} // Generic failure
				return (oWsh.Run("%comspec% /c nslookup -type=TXT " + sOwner + " " + sServer + " | find /i \"Non-existent\" >nul",__HIO.hide,true) != 0);	
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
		finally{
			
		}
	}

	this.atype = function atype(sOpt,sContainerName,sDomainName,sServer,sIPaddress,sOwner){
		try{
			sDomainName = typeof(sDomainName) == "string" ? sDomainName : sContainerName
			var sClass = "MicrosoftDNS_AType"
			if(sOpt == "getatypes"){			
				sQuery = "Select ContainerName,DescriptiveText,DnsServerName,DomainName,OwnerName,RecordData from " + sClass + " where ContainerName='" + sContainerName + "' And DomainName='" + sDomainName + "'"
				return wmi_win32_class("simple",sClass,sQuery,"\n",server,namespace,username,password);
			}
			else if(sOpt == "getatypesall"){
				if(aDomains = this.getDomains(sContainerName,sDomainName)){
					var aOClass = []
					for(var i = 0, len = aDomains.length; i < len; i++){
						if(aClass = this.domain("gettxttypes",sContainerName,aDomains[i])){
							aOClass.push(aClass)
						}
						__HUtil.kill(aClass)
					}
					__HUtil.kill(aDomains)
					return aOClass
				}			
			}
			else if(sOpt == "delatype"){
				var oClass = this.getclass(sClass)
				sOwner = typeof(sOwner) == "string" ? sOwner : sServer + "." + sContainerName
				var oResult = oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,this.RecordClass,this.TTL,sText)
			}
			else if(sOpt == "addatype"){
				var oClass = this.getclass(sClass)
				sOwner = !__H.isStringEmpty(sOwner) ? sOwner : sServer + "." + sContainerName
				try{
					oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,this.RecordClass,this.TTL,sIPaddress);
				}
				catch(ee){} // Generic failure
				return (oWsh.Run("%comspec% /c nslookup -type=A " + sOwner + " " + sServer + " | find /i \"Non-existent\" >nul",__HIO.hide,true) != 0);
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.cnametype = function cnametype(sOpt,sContainerName,sDomainName,sServer,sOwner,sPrimaryName){
		try{
			sDomainName = typeof(sDomainName) == "string" ? sDomainName : sContainerName
			var oClass
			var sClass = "MicrosoftDNS_CNAMEType"
			if(sOpt == "getcnametypes"){
				sQuery = "Select ContainerName,DnsServerName,DomainName,OwnerName,PrimaryName,RecordClass,RecordData,TextRepresentation from " + sClass + " where ContainerName='" + sContainerName + "' And DomainName='" + sDomainName + "'"
				/*
					Caption: null
					ContainerName: Nosliw.LAN
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
				return wmi_win32_class("simple",sClass,sQuery,"\n",server,namespace,username,password);
			}
			else if(sOpt == "getcnametypesall"){			
				if(aDomains = this.getDomains(sContainerName,sDomainName)){
					var aOClass = []
					for(var i = 0, len = aDomains.length; i < len; i++){
						if(aClass = this.domain("getcnametypes",sContainerName,aDomains[i])){
							aOClass.push(aClass)
							aClass = null
						}
					}
					return aOClass
				}
			}		
			else if(sOpt == "delcname"){
				if(oClass = this.getclass(sClass)){
					sOwner = typeof(sOwner) == "string" ? sOwner : sServer + "." + sContainerName
					var oResult = oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,this.RecordClass,this.TTL,sText)
				}			
			}
			else if(sOpt == "addcname"){
				if(oClass = this.getclass(sClass)){
					var oRe = new RegExp("[a-z0-9\-_]+\." + sDomainName + "$","ig"); // Must have ".Domain" 
					sOwner = sOwner.isSearch(oRe) ? sOwner : sOwner + "." + sDomainName // local-printer-server.Domain
					try{
						oClass.CreateInstanceFromPropertyData(sServer,sContainerName,sOwner,this.RecordClass,this.TTL,sPrimaryName);
					}
					catch(ee){} // Generic failure
					return (oWsh.Run("%comspec% /c nslookup -type=CNAME " + sOwner + " " + sServer + " | find /i \"Non-existent\" >nul",__HIO.hide,true) != 0);
				}
			}
			else if(sOpt == "setcname"){
				/*
				if(oClass = this.getclass(sClass)){
					sOwner = typeof(sOwner) == "string" ? sOwner : sServer + "." + sContainerName
					oClass.Modify(this.TTL,sPrimaryName)
				}
				*/
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.soatype = function soatype(sOpt,sContainerName,sDomainName,sServer,sOwner,sPrimaryName){
		try{
			sDomainName = typeof(sDomainName) == "string" ? sDomainName : sContainerName
			if(sOpt == "getserial"){
				var sQuery = "Select SerialNumber from MicrosoftDNS_SOAType where ContainerName='" + sContainerName + "'"
				var oColItems = o_service.ExecQuery(sQuery,null,48);
				for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !__H.$stopit; oEnum.moveNext()){
					var oItem = oEnum.item();
					return oItem.SerialNumber()
				}
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

})

