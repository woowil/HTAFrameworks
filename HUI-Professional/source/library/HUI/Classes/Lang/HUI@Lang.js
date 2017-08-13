// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

//(function(){
	
	//if(__H.d_registered.Exists("Lang")){return}
	
	__H.register(__H,"Lang","A Language Class",function Lang(){
		var o_this = this
		var b_initialized = false
		
		
		/////////////////////////////////////
		//// DEFAULT
		
		var o_options = {
			
		}
		
		var initialize = function initialize(bForce){
			if(b_initialized && !bForce) return;
			
			d_langs = new ActiveXObject("Scripting.Dictionary")
			o_Lang  = {}
			
			b_initialized = true
		}
		
		this.setOptions = function setOptions(oOptions){
			if(__H.isObject(oOptions)) return Object.extend(o_options,oOptions,true)
			return false
		}
		
		this.getOptions = function(){
			return o_options
		}
		
		this.disclaimer = "This sample code is provided AS IS WITHOUT " +
		"WARRANTY OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY " +
		"OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code " +
		"is to be undertaken entirely at your risk, and the " +
		"results that may be obtained from it are dependent on the user. " +
		"Please note to fully back up files and system(s) on a regular " +
		"basis. A failure to do so can result in loss of data or damage to systems. "
		
		/////////////////////////////////////
		//// 	
		
		this.UTF8 = {
			encode: function(s){
				for(var c, i = -1, l = (s = s.split("")).length, o = String.fromCharCode; ++i < l;
     			s[i] = (c = s[i].charCodeAt(0)) >= 127 ? o(0xc0 | (c >>> 6)) + o(0x80 | (c & 0x3f)) : s[i]
				)
				return s.join("");
			},
			
			decode: function(s){
				for(var a, b, i = -1, l = (s = s.split("")).length, o = String.fromCharCode, c = "charCodeAt"; ++i < l;
    			((a = s[i][c](0)) & 0x80) &&
    			(s[i] = (a & 0xfc) == 0xc0 && ((b = s[i + 1][c](0)) & 0xc0) == 0x80 ?
    			o(((a & 0x03) << 6) + (b & 0x3f)) : o(128), s[++i] = "")
				);
				return s.join("");
			}
		};
		
		
	})
//}());

var __HLang = new __H.Lang()


