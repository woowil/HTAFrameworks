// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@Sys@Mgmt.js")

__H.register(__H.Sys.Mgmt,"Service","System Registry",function Service(){
	var o_service = null
	var d_services = new ActiveXObject("Scripting.Dictionary")
		
	this.initialize = function initialize(sComputer,sUser,sPass,bForce){
		try{
			if(__HWMI.setServiceWMI(sComputer,"root\\cimv2",sUser,sPass)){
				o_service = __HWMI.setServiceWMI(bForce)
			}
			return o_service
		}
		catch(ee){			
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getDependents = function getDependents(sService){ // protected function
		var sDict = this.computer + ":" + sService
		if(d_services.Exists(sDict)) return d_services(sDict)

		var sFile = __HIO.temp + "\\" + oFso.GetTempName()
		var aServices = [];
		var sCmd = "%comspec% /c sc \\\\" + this.computer + " enumDepend \"" + sService + "\" | find /i \"SERVICE_NAME:\">" + sFile

		__HLog.log("## Getting service dependentcies for: " + sService)

		if(oWsh.Run(sCmd,$Env.hide,true) == 0){
			var oFile = oFso.OpenTextFile(sFile,$Env.read,false,$Env.TristateUseDefault)
			var oRe = new RegExp("SERVICE_NAME: ([a-z0-9]+)","ig")
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				if(sLine.match(oRe)) aServices.push(RegExp.$1)
			}
			oFile.close()
		}
		(oFso.GetFile(sFile)).Delete(); // Also deletes empty files

		d_services.Add(sDict,aServices)
		return aServices
	}

	this.getDescription = function getDescription(sService){
		var sFile = __HIO.temp + "\\" + oFso.GetTempName()
		var sDesc = ""
		var sCmd = "%comspec% /c sc \\\\" + this.computer + " qdescription \"" + sService + "\" | find /i \"DESCRIPTION\">" + sFile

		__HLog.log("## Getting service description for: " + sService)

		if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
			var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault)
			var oRe = new RegExp("[ \t]*DESCRIPTION[ \t]*:[ \t]*(.+)\n$","ig")
			var sLine = oFile.ReadAll()
			oFile.close()
			if(sLine.match(oRe)) sDesc = RegExp.$1
		}
		__HLog.log("### " + sDesc)
		return sDesc
	}

	this.stop = function stop(sService){ // protected function
		if(__H.isString(sService)){
			var a = this.getDependents(sService)
			a = a.concat(sService) // put the main service last
			for(var i = 0, iLen = a.length; i < iLen; i++){
				__HLog.log("### Stopping service: " + a[i])
				if(oWsh.Run("%comspec% /c sc \\\\" + this.computer + " stop \"" + a[i] + "\" >nul",__HIO.hide,true) != 0){
					__HLog.log("#### Either already stopped, disabled or unable to stop service: " + a[i] + " on computer: " + this.computer)
					return false
				}
			}
			return true
		}
		return false
	}

	this.start = function start(sService){ // protected function 
		if(__H.isString(sService)){
			var aServices = new Array(sService) // put the main service first
			aServices = aServices.concat(this.getDependents(sService))
			for(var i = 0, iLen = aServices.length; i < iLen; i++){
				__HLog.log("### Starting service: " + aServices[i])
				if(oWsh.Run("%comspec% /c sc \\\\" + this.computer + " start \"" + aServices[i] + "\" >nul",__HIO.hide,true) != 0){
					__HLog.log("#### Either already started or unable to start service: " + aServices[i] + " on computer: " + this.computer)
					return false
				}
			}
			return true
		}
		return false
	}

	this.restart = function restart(sService){ // protected function
		return (this.stop(sService) && this.start(sService))
	}
	
	this.restartRPC = function restartRPC(sService){ // protected function
		// The "Computer Browser" service is depended on the "server" service. Stopping "server" will also stop Computer Browser.  Stopping "workstation" will disconect all mappings
		return (this.restart("netlogon") && this.restart("browser"))
	}
	
	this.getInfo = function getInfo(sService){
		if(__H.isString(sService)){
			__HLog.log("### Getting information on service: " + sService)
			if(!o_service && !this.initialize()) return false
			var oColItems = o_service.ExecQuery("Select Description,DisplayName,Name,PathName,ProcessId,Started,StartName,StartMode,State,StartName,Status from Win32_Service where Name='" + sService + "' OR DisplayName='" + sService + "'","WQL",48)
			var o = {}
			for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd() && !__H.$stopit; oEnum.moveNext()){
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
			__HUtil.kill(oEnum,oColItems)
			return o
		}
		return false
	}
})

var __HService = new __H.Sys.Mgmt.Service()