// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@IO@File.js")

__H.register(__H.IO.File,"JSNormalize","JSNormalize",function JSNormalize(){
	var o_this = this
	var b_initialized = false
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		__HLog.debug("Initializing class: __H.IO.File.JSNormalize")
		
		
		b_initialized = true
	}
	initialize()
	
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
})

var __HJSNormalize = new __H.IO.File.JSNormalize()