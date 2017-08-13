// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.

//**Start Encode**

__H.include("HUI@IO@File.js")

__H.register(__H.IO.File,"Wrapper","A Wrapper Class",function Wrapper(){
	var o_this = this
	var b_initialized = false
	
	var d_wrappers
	var o_template
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		d_wrappers = new ActiveXObject("Scripting.Dictionary")
		o_template  = {}
		
		b_initialized = true
	}
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(o_options,oOptions,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	/////////////////////////////////////
	//// 

	this.isWrapper = function isWrapper(sArgument){
		return true
	}

	var loadWrapper = function loadWrapper(sArgument){
		try{
			initialize()
			 if(o_this.isWrapper(sArgument)){
				
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}		
	}

	this.setWrapper = function setWrapper(sArgument){
		try{
			this.reset()
			if(!load(sArgument)){
				__HExp.ArgumentIllegal("Argument[0] is not a valid file: " + sArgument)
			}
			
			o_template.value = "something"
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getWrapper = function getWrapper(){
		return o_template.value
	}
	
	this.toSomething = function toSomething(){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.reset = function reset(){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	/////////////////////////////////////
	//// MAIN
	
	var pstools = function(sFile){
		if(!oFso.FileExists(sFile)){
			__HLog.debug("Unable to locate PSTOOLS file: " + sFile)
			return false
		}
		else if(!o_this.ping()) return false
		return !o_this.isStringEmpty(o_this.domainuser,o_this.password) ? " -u " + o_this.domainuser + " -p " + o_this.password : "";
	}

	this.psshutdown = function psshutdown(sFile,bShutdown,bLogoff){
		try{
			var s = ""
			if(!(s=pstools(sFile))) return false;

			if(bShutdown) s += " -f -c -t 360 -m \"System will be shutdown."
			else if(bLoggoff)  s += " -f -0 -t 90 -m \"Console user will be loggoff."
			else s += " -f -r -c -t 120 -m \"System will be rebooted."

			return (oWsh.Run("\"" + sFile + s + " \\\\" + __HSys.computer,__HIO.hide,true) == 300)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.psexec = function psexec(sFile,bShowExit){
		try{
			var s
			if(!(s=pstools(sFile))) return false;

			var sCmd = "%comspec% /k \"" + sFile + "\" \\\\" + __HSys.computer + s + " cmd" + (bShowExit ? " & exit" : "")
			return (oWsh.Run(sCmd,__HIO.show,true) == 300)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
})
	
var __HWrapper = new __H.IO.File.Wrapper()
	
