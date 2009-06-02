// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-com-wmi-network.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-01
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

function wmi_network_class(bStatus){
	try{
		var service = false		
		var macaddress = null
		var properties = null
		var POPUP_WAIT = 15		
		var status = bStatus ? true : false
		
		// Initial
		var oReg = new Object()
		oReg.read = 1
		oReg.hide = 0
		oReg.write = 2
		oReg.TristateUseDefault = -2
		oReg.temp = oFso.GetSpecialFolder(2)
		
		var getProperties = function(sClass,oClassItem,o,WQL){
			var p = wmi_class_property("getobject_simple",service,sClass)
			o = o ? o : new Object()
			sClass = new String(sClass)
			for(var i = 0, len = p.length; i < len; i++){
				var oRe = new RegExp(".+" + new String(p[i].name) + ".+","g")
				if(WQL.match(oRe)){
					var oo = oClassItem[p[i].name], v
					if(p[i].isdatetime) oo = wmi_class_date(oo)
					else if(p[i].isarray) oo = ((v=wmi_class_vbarray(oo,p[i].name,"\n")) ? v.stream : " ").replace(/.+:[ \t]+(.+)/gm,"$1")
					o[p[i].name] = oo
				}
			}
			return o
		}
		
		this.setService = function(oService,sComputer,sUser,sPass){
			if(oService) service = oService
			else {
				sComputer = sComputer ? sComputer : oWno.ComputerName
				sUser = sUser ? sUser : null
				sPass = sPass ? sPass : null
				service = wmi_wbem_service(sComputer,"root\\Cimv2",sUser,sPass);
			}
			return service
		}
		
		this.getQuery = function(WQL,sComputer){
			service = service ? service : this.setService(null,sComputer)
			return service.ExecQuery(WQL,null,48)
		}
		
		this.getMAC = function(sComputer){
			if(macaddress) return macaddress
			var sFile = oReg.temp + "\\getMAC.log"
			var sCmd = "%comspec% /c nbtstat -a " + sComputer + " | find /i \"MAC ad\" > " + sFile;
			if(oWsh.Run(sCmd,oReg.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault)
				var s = oFile.ReadAll() // MAC Address = 00-0D-56-FE-C1-DE
				oFile.close()
				if(s.match(/[a-z \t]+=[ ]([a-z0-9\-]+)[ ]*/ig)){
					macaddress = (RegExp.$1).replace(/[\-]/ig,":")
				}			
			}
			if(oFso.FileExists(sFile)){
				(oFso.GetFile(sFile)).Delete(); // deletes also empty files
			}
			return macaddress
		}
		
		this.setWINS = function(sComputer,sPrimary,sSecondary){
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.getMAC(sComputer) + "'"
			for(var oItem, oEnum = new Enumerator(this.getQuery(WQL,sComputer)); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext()){
				oItem = oEnum.item()
				oItem.SetWINSServer(sPrimary,(sSecondary ? sSecondary : null))
				return true
			}
			return false
		}
		
		this.setDNS = function(sComputer,sPrimary,sSecondary){
			var WQL = "select MACAddress from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.getMAC(sComputer) + "'"
			var aVBDNS = this.JSVBArray(sPrimary,(sSecondary ? sSecondary : null))
			for(var oItem, oEnum = new Enumerator(this.getQuery(WQL,sComputer)); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext()){
				oItem = oEnum.item()
				oItem.SetDNSServerSearchOrder(aVBDNS)
				return true
			}
			return false
		}
		
		this.getIPConfigAll = function(sComputer){
			var WQL = "select DefaultIPGateway,DHCPEnabled,DHCPLeaseExpires,DHCPLeaseObtained,DNSDomain,DNSDomainSuffixSearchOrder,DNSServerSearchOrder,FullDNSRegistrationEnabled,IPAddress,IPSubnet,MACAddress,ServiceName,WINSPrimaryServer,WINSSecondaryServer from Win32_NetworkAdapterConfiguration where IPEnabled='true' and MACAddress='" + this.getMAC(sComputer) + "'"
			for(var oItem, o, oEnum = new Enumerator(this.getQuery(WQL,sComputer)); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext()){
				oItem = oEnum.item()
				o = getProperties("Win32_NetworkAdapterConfiguration",oItem,null,WQL)
			}			
			WQL = "select DeviceID,Manufacturer,NetConnectionID,ProductName,TimeOfLastReset from Win32_NetworkAdapter where MACAddress='" + this.getMAC(sComputer) + "'"
			for(var oItem, oEnum = new Enumerator(this.getQuery(WQL,sComputer)); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext()){
				oItem = oEnum.item()
				o = getProperties("Win32_NetworkAdapter",oItem,o,WQL)
			}
			return o
		}
		
		this.objectToHTML = function(o,bHead){
			var h = '<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">'
			//for(oo in o) h = h + '<th>' + oo + '</th>'
			//h = h + '</tr>'
			for(oo in o){
				h = h + '\n<tr><td>' + oo + '</td><td>' + o[oo] + '</td></tr>'
			}
			//alert(h)
			return h + '\n</table>'
		}
		
		this.objectToCSV = function(o){
			var c = ""
			for(oo in o){
				c = c + o[oo] + ';'
			}
			return c.replace(/(.+);$/g,"$1")
		}
		
		this.reboot = function(sComputer){
			var oColItems = GetObject("winmgmts:{(RemoteShutdown)}//" + sComputer + "/root/cimv2").ExecQuery("Select * from Win32_OperatingSystem where Primary=True")
			for(var oItem, oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext()){
				oItem = oEnum.item()
				oItem.Reboot()
			}
		}
		
		this.JSVBArray = function(){
			var d = new ActiveXObject("Scripting.Dictionary")
			for(var i = 0, len = arguments.length; i < len; i++){
				d.Add("d" + i,arguments[i])
			}
			return d.Items()
		}
		
		this.echo = function(sMessage){
			if(status){
				try{
					js_log_print("log_result",sMessage)
				}
				catch(ee){
					WScript.Echo(sMessage)
				}				
			}
		}
		
		this.popup = function(sMsg){
			try{
				alert(sMsg)
			}
			catch(ee){
				oWsh.Popup(sMsg,POPUP_WAIT,"Network Class",32 + 1)
			}
		}
	}
	catch(e){
		wmi_log_error(2,e);
		return false;
	}
}

