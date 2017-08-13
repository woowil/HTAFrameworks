// Copyright© 2003-2013. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.include("HUI@Lang.js")

//(function(){
__H.register(__H.Lang,"Locales","Locales",function Locales(){
	var o_this = this
	var b_initialized = false
	
	var d_Localess
	var o_Locales
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		say : "Hello World"
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		__HLog.debug("Initializing class: __H.Parent.Class.Locales")
		
		d_Localess = new ActiveXObject("Scripting.Dictionary")
		o_Locales  = {}
		
		b_initialized = true
	}
	
	this.setOptions = function setOptions(oOptions){
		if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
		return false
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	this.toString = function(){
		return this //.getLocaless().toString()
	}
	
	/////////////////////////////////////
	//// 

	this.isLocales = function isLocales(sArgument){
		return true
	}

	var loadLocales = function loadLocales(sArgument){
		try{
			initialize()
			 if(o_this.isLocales(sArgument)){
				
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}		
	}

	this.setLocales = function setLocales(sLocales){
		try{
			this.reset()
			if(!loadLocales(sLocales)){
				__HExp.ArgumentIllegal("Argument[0] is invalid for: " + sLocales)
			}
			
			o_Locales.value = "something"
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getLocales = function getLocales(){
		return o_Locales.value
	}
	
	this.getLocaless = function getLocaless(){
		this.loadLocales()
		return new VBArray(d_Localess.Keys()).toArray()
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
	
	
	
})
//}());
	
var __HLocales = new __H.Lang.Locales()
	
