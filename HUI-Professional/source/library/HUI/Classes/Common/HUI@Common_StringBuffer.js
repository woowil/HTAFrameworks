// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.

//**Start Encode**

__H.include("HUI@Common.js")

//(function(){  // can't use this here because this class is always called as an instance
__H.register(__H.Common,"StringBuffer","StringBuffer Class",function StringBuffer(s){
	var o_this = this

	/////////////////////////////////////
	////

	// http://www.softwaresecretweapons.com/jspwiki/javascriptstringconcatenation
	var a_buffer = []
	var i_length = 0 // Prevents manipulation of this.length

	this.length  = 0

	this.toString = function(delim){
		// DO NOT REMOVE OR CHANGE THIS FUNCTION!!
		return a_buffer.join(typeof(delim) == "string" ? delim : "");
	}

	this.toJSON = function(){
		return this.toString()
	}

	this.valueOf = function(){
		// DO NOT REMOVE OR CHANGE !!
		// 1. valueOf is used by call and apply built-in functions
		// 2. valueOf is used when concatenating using the "+, +=" operators
		// 3. valueOf is used when comparing using the "<, <=, >=, ==" operators
		return this.toString()
	}

	this.append = function(s,bInFront,bReset){
		if(typeof(s) == "string"){
			if(bReset) this.empty() //  || this.length == 0 // BUG!! WILL CREATE STRANGE FAILURE... append => undefined
			!bInFront ? a_buffer.push(s) : a_buffer.splice(0,0,s)
			i_length += s.length
			this.length = i_length
			delete s
		}
		return this
	}
	
	this.append(s)

	this.appendLine = function(s,bInFront,bReset){
		if(typeof(s) != "string") s = ""
		return this.append(s.concat("\r\n"),bInFront,bReset)
	}

	this.isEmpty = function(){
		return a_buffer.length == 0
	}

	this.empty = function(){
		a_buffer.empty()
		this.length = i_length = 0
		return this
	}

	this.compact = function(){
		return this.trim(true)
	}

	this.clear = function(){
		return this.empty()
	}

	this.concat = function(s){
		if(typeof(s) != "string") return this.toString()
		var a = []
		a.push(this.toString(),s)
		return a.join("")
		//return this.toString().concat(s)
	}

	this.trim = function(bInner){
		var s = this.toString().trim(bInner)
		this.empty()
		return this.append(s).toString()
	}

	this.match = function(oRe){
		if(oRe instanceof RegExp){
			return this.toString().match(oRe)
		}
		return null
	}

	this.search = function(oRe){
		if(oRe instanceof RegExp){
			return this.toString().search(oRe)
		}
		return -1
	}

	this.substring = function(i,j){
		return this.toString().substring(i,j)
	}

	this.substr = function(i){
		return this.toString().substring(i,i_length)
	}

	this.split = function(s){
		return this.toString().split(s)
	}

	this.shuffle = function(){
		var s = this.toString().shuffle()
		this.empty()
		return this.append(s).toString()
	}

	this.encode = function(){
		var s = this.toString().encode()
		this.empty()
		return this.append(s).toString()
	}

	this.random = function(){
		return this.toString().charAt((0).random(i_length))
	}
})

var __HStringBuffer = new __H.Common.StringBuffer()
var __HBuf          = new __H.Common.StringBuffer()
