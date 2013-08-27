// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@Util.js")

// This is similar to HASH functions
__H.register(__H.Util,"Dictionary","Dictionary ActiveXObject",function Dictionary(){
	var o_this = this
	var b_initialized = false
	
	var d_dictionary
	var o_dictionary
	
	this.Count = 0
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		__HLog.debug("Initializing class: __H.Util.Dictionary")
		
		d_dictionary = new ActiveXObject("Scripting.Dictionary")
		
		
		o_dictionary  = {}
		
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
	
	this.toString = function(){
		return this.toArrayKeys().toString()
	}
	
	this.valueOf = function(){
		return this.getDictionary()
	}
	
	this.isDictionary = function isDictionary(sArgument){
		return true
	}
	
	var loadDictionary = function loadDictionary(sArgument){
		try{
			initialize()
			 if(o_this.isDictionary(sArgument)){
				
				return true
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}		
	}

	this.setDictionary = function setDictionary(sArgument){
		try{
			this.reset()
			if(!loadDictionary(sArgument)){
				__HExp.ArgumentIllegal("Argument [0] is not a valid: " + sArgument)
			}
			
			o_dictionary.value = "something"
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getDictionary = function getDictionary(){
		return d_dictionary
	}
	
	this.Add = function Add(sKey,sItem){
		if(!d_dictionary.Exists(sKey)){
			d_dictionary.Add(sKey,sItem)
			this.Count++
		}
		else __HLog.popup("Key: '" + sKey + "' already exists in Dictionary")
		return d_dictionary
	}
	
	this.Exists = function(sKey){
		return d_dictionary.Exists(sKey)
	}
	
	this.Remove = function Remove(sKey){
		if(d_dictionary.Exists(sKey)){
			d_dictionary.Remove(sKey)
			this.Count--
		}
		return d_dictionary
	}
	
	this.RemoveAll = function(){
		this.Count = 0
		d_dictionary.RemoveAll()
		return d_dictionary
	}
	
	this.toArrayKeys = function(){
		return (new VBArray(d_dictionary.Keys())).toArray()
	}
	
	this.toArrayItems = function(){
		return (new VBArray(d_dictionary.Items())).toArray()
	}
	
	this.concat = function concat(d){
		try{
			var dd = this.clone()
			if(!this.isDictionary(d)) return dd
			
			var a = (new VBArray(d.Keys())).toArray()			
			for(var i in a){
				if(!dd.exists(a[i])){
					dd.Add(a[i],d(a[i]))
				}
			}
			a.empty()
			return dd
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.clone = function(){
		var d = new ActiveXObject("Scripting.Dictionary")
		var a = this.toArrayKeys()
		
		for(var i in a){
			d.Add(a[i],d_dictionary(a[i]))
		}
		a.empty()
		return d
	}
	
	this.reset = function reset(){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.empty = function(){
		return this.RemoveAll()
	}
	
	/////////////////////////////////////
	//// MAIN
	
	
	
})
	
var __HDictionary = __H.Util.Dictionary
	
