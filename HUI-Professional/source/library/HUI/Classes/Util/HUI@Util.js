// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.register(__H,"Util","Utilities",function Util(){
	var o_this = this
	var b_initialized = false
	
	// The ActiveX Control of AutoItX3.dll
	this.autoit = null	

	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;
		
		b_initialized = true
	}
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(oOptions,true)
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
	

	this.sleep = function sleep(iMilli){
		if(typeof(iMilli) != "number" && iMilli < 1) return false;
		else if(typeof(iMilli) == "string") iMilli = parseInt(iMilli);
		else return false;
		//if(__H.isFunction(this.autoit)) this.autoit.Sleep(iMilli);
		//else if(__H.isWScript()) WScript.Sleep(iMilli);
		return (oWsh.Run("%comspec% /c sleep -m " + iMilli + ">nul",__HIO.hide,true) == 0)
	}

	this.random = function random(hval,lval){
		if(typeof hval != "number") return hval;
		if(typeof lval != "number") lval = 0;
		return Math.floor(Math.random()*hval + lval)
	}

	this.matrix2D = function(x,y,v){
		if(typeof x != "number" && x >= 0) x = 0;
		if(typeof y != "number" && y >= 0) y = 0;
		v = v ? v : "#"
		for(var i = 0, a = []; i < x; i++){
			a[i] = []
			for(var j = 0; j < y; j++){
				a[i][j] = v
			}
		}
		return a
	}

	this.kill = function(){
		try{
			var oRe1 = /boolean|unknown|string|undefined|number/g
			var oRe2 = /object|function/g
			for(var i = arguments.length; i; i--){
				if(!arguments[i-1]) continue
				else if(typeof(arguments[i-1]).isSearch(oRe1));
				else if(typeof(arguments[i-1]).isSearch(oRe2) && typeof(arguments[i-1].empty) == "function"){
					arguments[i-1].empty()
				}
				delete arguments[i-1] // delete: http://users.adelphia.net/~daansweris/js/special_operators.html
			}
		}
		catch(e){}
	}
})

var __HUtil	= new __H.Util()