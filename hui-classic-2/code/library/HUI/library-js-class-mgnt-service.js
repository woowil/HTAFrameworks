// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-com-wmi-network.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-01
Version:
Dependency:  library-js.js, library-js__prototype.js

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

// Function _Class() must be loaded
class_mgnt_service.prototype = new _Class("Service","Class for managing services")

function class_mgnt_service(oService,sComputer,sUser,sPass,bInit,bStatus,bDebug){
	try{
		this.wmi_cimv2 = oService
		this.services = new ActiveXObject("Scripting.Dictionary")
		this.computer = typeof(sComputer) == "string" ? sComputer : oWno.ComputerName
		this.username = typeof(sUser) == "string" ? sUser : null
		this.password = typeof(sPass) == "string" ? sPass : null
		this.isStatus = bStatus ? true : false
		this.isDebug = bDebug ? true : false
					
		this.getDependents = function(sService){ // protected function
			var sDict = this.computer + ":" + sService
			if(this.services.Exists(sDict)) return this.services(sDict)

			var sFile = this.fso.temp + "\\Service_getDependents.tmp"
			var aServices = new Array();
			var sCmd = "%comspec% /c sc \\\\" + this.computer + " enumDepend \"" + sService + "\" | find /i \"SERVICE_NAME:\">" + sFile

			this.echo("## Getting service dependentcies for: " + sService)

			if(oWsh.Run(sCmd,this.fso.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,this.fso.read,false,this.fso.TristateUseDefault)
				var oRe = new RegExp("SERVICE_NAME: ([a-z0-9]+)","ig")
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(sLine.match(oRe)) aServices.push(RegExp.$1)
				}
				oFile.close()
			}
			(oFso.GetFile(sFile)).Delete(); // Also deletes empty files
	
			this.services.Add(sDict,aServices)
			return aServices
		}

		this.getDescription = function(sService){
			var sFile = this.fso.temp + "\\Service_getDescription.tmp"
			var sDesc = ""
			var sCmd = "%comspec% /c sc \\\\" + this.computer + " qdescription \"" + sService + "\" | find /i \"DESCRIPTION\">" + sFile

			this.echo("## Getting service description for: " + sService)

			if(oWsh.Run(sCmd,this.fso.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,this.fso.read,false,this.fso.TristateUseDefault)
				var oRe = new RegExp("[ \t]*DESCRIPTION[ \t]*:[ \t]*(.+)\n$","ig")
				var sLine = oFile.ReadAll()
				oFile.close()
				if(sLine.match(oRe)) sDesc = RegExp.$1
			}
			this.echo("### " + sDesc)
			return sDesc
		}
	
		this.stop = function(sService){ // protected function
			if(this.isString(sService)){
				var aServices = this.getDependents(sService)
				aServices = aServices.concat(sService) // put the main service last
				for(var i = 0, len = aServices.length; i < len; i++){
					this.echo("### Stopping service: " + aServices[i])
					if(oWsh.Run("%comspec% /c sc \\\\" + this.computer + " stop \"" + aServices[i] + "\" >nul",this.fso.hide,true) != 0){
						this.echo("#### Either already stopped, disabled or unable to stop service: " + aServices[i] + " on computer: " + this.computer)
						return false
					}
				}
				return true
			}
			return false
		}

		this.start = function(sService){ // protected function 
			if(this.isString(sService)){
				var aServices = new Array(sService) // put the main service first
				aServices = aServices.concat(this.getDependents(sService))
				for(var i = 0, len = aServices.length; i < len; i++){
					this.echo("### Starting service: " + aServices[i])
					if(oWsh.Run("%comspec% /c sc \\\\" + this.computer + " start \"" + aServices[i] + "\" >nul",this.fso.hide,true) != 0){
						this.echo("#### Either already started or unable to start service: " + aServices[i] + " on computer: " + this.computer)
						return false
					}
				}
				return true
			}
			return false
		}

		this.restart = function(sService){ // protected function
			return (this.stop(sService) && this.start(sService))
		}
		
		this.restartRPC = function(sService){ // protected function
			// The "Computer Browser" service is depended on the "server" service. Stopping "server" will also stop Computer Browser.  Stopping "workstation" will disconect all mappings
			return (this.restart("netlogon") && this.restart("browser"))
		}
		
		this.getInfo = function(sService){
			if(this.isString(sService)){
				this.echo("### Getting information on service: " + sService)
				if(!this.setService()) return false
				var oColItems = this.wmi_cimv2.ExecQuery("Select Description,DisplayName,Name,PathName,ProcessId,Started,StartName,StartMode,State,StartName,Status from Win32_Service where Name='" + sService + "' OR DisplayName='" + sService + "'","WQL",48)
				var o = new Object()
				for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !js.pro.stopit; oEnum.moveNext()){
					var oItem = oEnum.item();
					o.Description = oItem.Description
					o.DisplayName = oItem.DisplayName
					o.Name = oItem.Name
					o.PathName = oItem.PathName
					o.ProcessId = oItem.ProcessId
					o.Started = oItem.Started // true
					o.StartMode = oItem.StartMode // auto, manual, disabled
					o.State = oItem.State // Running
					o.StartName = oItem.StartName // LocalSystem
					o.Status = oItem.Status // OK
				}
				this.kill(oEnum,oColItems)
				return o
			}
			return false
		}		
		
		if(bInit) this.init(oService)
	}
	catch(e){
		js_log_error(2,e);
	}
}

