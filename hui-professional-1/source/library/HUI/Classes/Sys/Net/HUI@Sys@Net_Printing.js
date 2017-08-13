// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@Sys@Net.js")

__H.register(__H.Sys.Net,"Printing","Printing",function Printing(){
	var o_this = this
	var b_initialized = false
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		__HLog.debug("Initializing class: __H.Sys.Net.Printing")
		
		
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

var __HPrinting = new __H.Sys.Net.Printing()