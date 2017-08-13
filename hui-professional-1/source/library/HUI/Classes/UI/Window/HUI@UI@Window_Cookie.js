// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window.js")

__H.register(__H.UI.Window,"Cookie","HTML Cookie",function Cookie(){
	var name = "unknown"
	
	this.setName = function setName(sName){
		name = sName
	}
	
	this.getName = function getName(){
		return name
	}
	
	this.set = function set(value,expire){
		document.cookie = name + "=" + escape(value) + ((expire == null) ? "" : ("; expires=" + expire.toGMTString()));		
	}
	
	this.get = function get(){
	   var search = this.getName() + "=";
	   if(document.cookie.length > 0){ // if there are any cookies
	      var offset = document.cookie.indexOf(search)
	      if(offset != -1) { // if cookie exists
	         offset = offset + search.length
	         // set index of beginning of value
	         var end = document.cookie.indexOf(";", offset)
	         // set index of end of cookie value
	         if(end == -1) end = document.cookie.length
	         return unescape(document.cookie.substring(offset,end))
	      }
	   }
	}

	this.reg = function reg() {
	   var today = new Date()
	   var expires = (new Date()).setTime(today.getTime() + 1000*60*60*24*365)
	   this.set("TheCoolJavaScriptPage",this.getName(),expires)
	}

	this.save = function save(value,days){
		if(!isNaN(days)){
			var d = new Date();
			d.setTime(d.getTime() + (days*24*60*60*1000))
			var expires = "; expires=" + d.toGMTString()
		}
		else expires = "";
		document.cookie = this.getName() + "=" + value + expires + "; path=/";
	}

	this.read = function read(){
		var nameEQ = this.getName() + "="
		var a = document.cookie.split(';')
		for(var i = 0, iLen = a.length; i < iLen; i++){
			var c = a[i];
			while(c.charAt(0) == ' ') c = c.substring(1,c.length);
			if(c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length)
		}
		return false;
	}

	this.del = function del() {
		this.save("",-1)
		this.setName(null)
	}
})


function getCookie( name ) {
	var start = document.cookie.indexOf( name + "=" );
	var len = start + name.length + 1;
	if ( ( !start ) && ( name != document.cookie.substring( 0, name.length ) ) ) {
		return null;
	}
	if ( start == -1 ) return null;
	var end = document.cookie.indexOf( ';', len );
	if ( end == -1 ) end = document.cookie.length;
	return unescape( document.cookie.substring( len, end ) );
}

function setCookie( name, value, expires, path, domain, secure ) {
	var today = new Date();
	today.setTime( today.getTime() );
	if ( expires ) {
		expires = expires * 1000 * 60 * 60 * 24;
	}
	var expires_date = new Date( today.getTime() + (expires) );
	document.cookie = name+'='+escape( value ) +
		( ( expires ) ? ';expires='+expires_date.toGMTString() : '' ) + //expires.toGMTString()
		( ( path ) ? ';path=' + path : '' ) +
		( ( domain ) ? ';domain=' + domain : '' ) +
		( ( secure ) ? ';secure' : '' );
}

function deleteCookie( name, path, domain ) {
	if ( getCookie( name ) ) document.cookie = name + '=' +
			( ( path ) ? ';path=' + path : '') +
			( ( domain ) ? ';domain=' + domain : '' ) +
			';expires=Thu, 01-Jan-1970 00:00:01 GMT';
}