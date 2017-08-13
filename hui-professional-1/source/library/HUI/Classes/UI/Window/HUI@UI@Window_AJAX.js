// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.include("HUI@UI@Window.js")

__H.register(__H.UI.Window,"AJAX","AJAX",function AJAX(){
	this.active_reguest = 0
	var transport = null
	
	var options = {
		method			: 'post',
		asynchronous	: true,
		contentType		: 'application/x-www-form-urlencoded',
		encoding		: 'UTF-8',
		parameters		: '',
		evalJSON		: true,
		evalJS			: true
	}
	var url = null
	var body = ""
	
	this.initialize = function(sUrl){
		transport = this.getTransport()
		this.setOptions()
		this.request(sUrl)
	}
	
	this.setTransport = function(){
		try{
			if(transport != null) return transport
			return new XMLHttpRequest()
		}
		catch(e){
			var a = ["Msxml2.XMLHTTP","Microsoft.XMLHTTP"]
			for(var i = 0, iLen = a.length; i < iLen; i++){
				try{
					return new ActiveXObject(a[i]);
				}
				catch(e){}
			}
			throw new Error(__HLog.errorCode("error"),"Unable to invoke XML HTTP Request");
		}
	}
	
	this.getTransport = function(){
		return transport
	}
	
	this.setOptions = function(o){
		Object.extendObject(options,o)
		options.method = options.method.toLowerCase()
	}
	
	this.getOptions = function(){
		return options
	}
	
	this.setUrl = function(sUrl){
		if(!__H.isStringEmpty(sUrl)){
			if(sUrl.isSearch(__HKeys.ore_url) || sUrl.isSearch(/\.xml/ig)){
				//url = sUrl.isSearch()
			}
		}
		throw new Error(__HLog.errorCode("error"),"argument 0 is not an url")
	}
	
	this.getUrl = function(){
		return url
	}
	
	this.reguest = function(sUrl){
		if(options.method == "get"){
			this.setUrl(sUrl + "$" + options.parameter);
		}		
		try{
			var r = this.getTransport()
			r.open(options.method.toUpperCase(),url,options.asynchronous)
			
			body = options.method == 'post' ? (options.postBody || options.parameter) : null;
      		r.send(this.body);
		}
		catch(e){
			
		}
	}
	
	this.onStateChange = function(){
		if(transport.readyState == 4 && transport.status == 200){
		  //transport
		  this.onStateChange = __H.emptyFn
		  
		}
	}
	
	this.setRequestHeaders = function(){
		var headers = {
		  'X-Requested-With': 'XMLHttpRequest',
		  'X-AJAX-Version': this.$version,
		  'Accept': 'text/javascript, text/html, application/xml, text/xml, */*'
		}
		var o = this.getOptions()
		if(o.method == 'post'){
			headers['Content-type'] = o.contentType +(o.encoding ? '; charset=' + o.encoding : '');			
		}		
		
		for(var name in headers){
			if(headers.hasOwnProperty(name)){
				transport.setRequestHeader(name,headers[name]);
			}
		}
  }

  this.success = function(){
    var status = this.getStatus();
    return !status || (status >= 200 && status < 300);
  }

  this.getStatus = function(){
    try {
      return transport.status || 0;
    }
    catch(e){ return 0}
  }
})

var __HAJAX = new __H.UI.Window.AJAX()