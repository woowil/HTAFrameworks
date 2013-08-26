// Copyright© 2003-2013. nOsliw Solutions. All rights reserved.
//**Start Encode**

__H.include("HUI@ParentClassA@CurrentClass.js","HUI@ParentClassB@MandatoryClass_Name.js")

//(function(){

__H.register(__H.ParentClassA.CurrentClass,"Template","A Template CurrentClass",function Template(){
	var o_this = this
	var b_initialized = false
	
	var d_templates
	var o_template
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		say : "Hello World"
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		__HLog.debug("Initializing CurrentClass: __H.ParentClassA.CurrentClass.Template")
		
		d_templates = new ActiveXObject("Scripting.Dictionary")
		o_template  = {}
		
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
		return this //.getTemplates().toString()
	}
	
	/////////////////////////////////////
	//// 

	this.isTemplate = function isTemplate(sTemplate){
		return true
	}

	var loadTemplate = function loadTemplate(sTemplate){
		try{
			initialize()
			 if(o_this.isTemplate(sTemplate)){
				
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}		
	}

	this.setTemplate = function setTemplate(sTemplate){
		try{
			this.reset()
			if(!loadTemplate(sTemplate)){
				__HExp.ArgumentIllegal("Argument[0] is invalid for: " + sTemplate)
			}
			
			o_template.value = "something"
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getTemplate = function getTemplate(){
		return o_template.value
	}
	
	this.getTemplates = function getTemplates(){
		this.loadTemplate()
		return new VBArray(d_templates.Keys()).toArray()
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
	
	this._template = function _template(){
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
	
var __HTemplate = new __H.ParentClassA.CurrentClass.Template()
	
