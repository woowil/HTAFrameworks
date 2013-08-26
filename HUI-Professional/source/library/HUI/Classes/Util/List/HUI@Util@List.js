// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

// http://factsite.co.uk/en/wikipedia/l/li/linked_list.html

__H.include("HUI@Util.js")

__H.register(__H.Util,"List","Data Structure",function List(){
	var o_this = this
		var b_initialized = false
		
		
		/////////////////////////////////////
		//// DEFAULT
		
		var o_options = {
			
		}
		
		var initialize = function initialize(bForce){
			if(b_initialized && !bForce) return;
			
			
			
			b_initialized = true
		}
		
		this.setOptions = function setOptions(oOptions){
			if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
			return false
		}
		
		this.getOptions = function(){
			return o_options
		}
})