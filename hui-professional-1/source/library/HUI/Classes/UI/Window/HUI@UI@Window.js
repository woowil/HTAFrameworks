// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@UI.js")

__H.register(__H.UI,"Window","HTML",function Window(){
	var nua = navigator.userAgent
	var nan = navigator.appName
	var version = parseFloat(navigator.appVersion)
	var ContextTag = null

	this.bIExplore = nua.indexOf("Microsoft Internet Explorer") > -1
	this.bNS = nua.indexOf("Netscape") > -1
	this.bOpera = nua.indexOf("Opera") > -1
	this.bSafari = nua.indexOf("Safari") > -1

	this.bIERunning = !!window.ActiveX
	this.bIE = nua.indexOf("MSIE") > -1
	this.bIE5 = nua.indexOf("MSIE 5") > -1
	this.bIE55 = nua.indexOf("MSIE 5.5") > -1
	this.bIE50 = nua.indexOf("MSIE 5.0") > -1 && !this.bIE55
	this.bIE6 = nua.indexOf("MSIE 6") > -1 && !this.bOpera
	this.bIE7 = nua.indexOf("MSIE 7") > -1 && !this.bOpera
	this.bIE8 = nua.indexOf("MSIE 8") > -1 && !this.bOpera
	
	this.bIEnew = this.bIE55 || this.bIE6 || this.bIE7 || this.bIE8

	this.visible = this.bNS ? 'show' : 'visible'
	this.bJE = navigator.javaEnabled()
	this.w3 = !!(__H.$ && __H.$C);
	
	this.div_html = null
	
	this.removeFrames = function removeFrames(){
		if(self.parent.frames.length != 0){
			self.parent.location = document.location;
		}
	}

	this.removeHtml = function removeHtml(sHtml){
		if(typeof(sHtml) != "string") return ""
		return (sHtml.replace(/<[0-9a-z="': ;\-()]*>/ig,"")).replace(/<\/[a-z ]*>/ig,"")
	}

	this.setContextTag = function setContextTag(oElement){
		ContextTag = oElement
	}

	this.resize = function resize(){
		if(this.bIE){
			if(document.body.offsetWidth != window.saveInnerWidth || document.body.offsetHeight != window.saveInnerHeight){
				// do something
			}
		}
		else if(this.bNS){
			if(saveInnerWidth < window.innerWidth || saveInnerWidth > window.innerWidth || saveInnerHeight > window.innerHeight || saveInnerHeight < window.innerHeight ){
				window.history.go(0);
			}
		}
	}

	var compatMode = function(){
		if(document.compatMode){
			switch(document.compatMode){
				case "BackCompat" :
					return 0;
				case "CSS1Compat" :
					return 1;
				case "QuirksMode" :
					return 0;
				default : break;
			}
		}
		else {
			if(__HWindow.bIE5) return 0;
			else if(__HWindow.bSafari) return 1
		}
		return 0;
	}

	this.pagemode = 0
	
	var getScroll = function(sOpt){
		var sOpt = sOpt == "left" ? "scrollLeft" : "scrollTop"
		switch(__HWindow.pagemode){
			case 0 :
				return document.body[sOpt];
			case 1 :
				if(document.documentElement && document.documentElement[sOpt] > 0){
					return document.documentElement[sOpt];
				}
				else {
					return document.body[sOpt];
				}
			default : return document.body[sOpt]
		}
	}

	this.getScrollLeft = function getScrollLeft(){ return getScroll("left")}
	this.getScrollTop = function getScrollTop(){ return getScroll("top")}

	var getClient = function(sOpt){
		var sOpt = sOpt == "height" ? "clientHeight" : "clientWidth"
		switch(__HWindow.pagemode){
			case 0 :
				return document.body[sOpt]
			case 1 :
				if(__HWindow.bSafari) return self.innerHeight
				else {
					if(!__HWindow.bOpera && document.documentElement && document.documentElement[sOpt] > 0){
						return document.documentElement[sOpt]
					}
					else return document.body[sOpt]
				}
		}
	}

	this.getClientWidth = function getClientWidth(){ return getClient("width")}
	this.getClientHeight = function getClientHeight(){ return getClient("height")}

	this.getCoordinate = function getCoordinate(e){
		var e = (e || window.event)
		var o = {
			X : this.bSafari ? e.clientX - this.getScrollLeft() : e.clientX,
			Y : this.bSafari ? e.clientY - this.getScrollTop() : e.clientY
		}
		return o
	}

	this.getClientX = function getClientX(e){
		return (this.getCoordinate(e)).X
	}

	this.getClientY = function getClientY(e){
		return (this.getCoordinate(e)).Y
	}

	this.setContext = function setContext(sOpt){
	// SEEEEEE! http://msconline.maconstate.edu/tutorials/JSDHTML/JSDHTML12/jsdhtml12-02.htm
		try{ // http://msdn.microsoft.com/workshop/author/dhtml/reference/commandids.asp
			// http://msdn2.microsoft.com/en-us/library/ms533049.aspx
			if(sOpt == "redo"){
				document.execCommand("Redo");
			}
			else if(sOpt == "undo"){
				document.execCommand("Undo");
			}
			else if(sOpt == "cut"){
				//if(c("isedit")){ // ContextRange = document.selection.createRange()
					document.execCommand("Cut");
					//ContextTag.focus()
				//}
			}
			else if(sOpt == "copy"){
				//if(ContextRange){ // ContextRange = document.selection.createRange()
					document.execCommand("Copy"); // selection.createRange()
				//}
			}
			else if(sOpt == "paste"){
				//document.execCommand("Paste",false,clipboardData.getData('Text'));
				//htm_dbg_object(ContextTag)
				if(this.setContext("isedit")) ContextTag.focus() // Must have this!!
				else document.body.focus()
				document.execCommand("Paste");
			}
			else if(sOpt == "delete"){
				if(this.setContext("isedit")){
					//ContextTag.focus() // Must have this!!
					document.execCommand("Delete");
				}
			}
			else if(sOpt == "append"){
				//var sCopied = clipboardData.getData('Text');
				//document.execCommand('Paste',false,sCopied);
			}
			else if(sOpt == "uppercase"){
				if(this.setContext("isedit")){ // ContextRange = document.selection.createRange()
					document.selection.createRange().execCommand("Cut");
					var sCopied = (clipboardData.getData('Text')).toUpperCase();
					clipboardData.setData('Text',sCopied)
					this.setContext("paste")
				}
			}
			else if(sOpt == "selectall"){
				document.execCommand("SelectAll");
			}
			else if(sOpt == "isedit"){ // Excellence!!
				return (typeof(ContextTag) == "object" && ContextTag != null && ContextTag.nodeType == 1 && ContextTag.isTextEdit && !ContextTag.disabled && !ContextTag.readonly)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getSelection = function getSelection(){
		if(window.getSelection){ // Mozilla 1.75*, Safari 1.3
			return window.getSelection();
		}
		else if(document.getSelection){ // Explorer 5.2 Mac, Mozilla 1.75*, Netscape, Opera
			return document.getSelection();
		}
		else if(document.selection){ // Explorer 5 Win, Explorer 6 Win
			return document.selection.createRange().text;
		}
		return "";
	}

	this.nextObject = function(oElement){
		if(!__H.isElement(oElement)) return null
		do o = oElement.nextSibling;
		while(o && o.nodeType != 1);
		return o;
	}

	this.previousObject = function(oElement){
		if(!__H.isElement(oElement)) return null
		do o = oElement.previousSibling;
		while(o && o.nodeType != 1);
		return o;
	}
	
	this.setInnerHTML = function setInnerHTML(oElement){
		//if(__H.isElementEditable(this.div_html)){
		this.div_html.innerHTML = ""
		this.div_html.innerHTML = this.elementToHTML(oElement)
		//}
	}
	
	this.insertAfter = function(parent,node,referenceNode){
		if(!__H.isElement(parent,node,referenceNode)) return null
		parent.insertBefore(node,referenceNode.nextSibling);
	}
	
	this.elementToHTML = function elementToHTML(oElement){
		try{
			if(!oElement) return "[null]"
			if(typeof(oElement) == "string") oElement = __H.byIds(oElement)
			/*
			if(typeof(oElement) == "object" && oElement.nodeType == 9); // window Object
			
			else if(!__H.isElement(oElement) && !__H.isElementArray(oElement)){
				__HLog.popup("Not a standard HTML DOM Element")
				return;
			}
			*/
			var bHas = typeof(oElement.hasOwnProperty) == "function"
			
			var sHtml = new __H.Common.StringBuffer()
			sHtml.append('\n<table bgcolor="#EEEEEE" cellspacing="1" cellpadding="2" align="center" bordercolor="#000000" class="mba-table-1">')
			sHtml.append('\n<tr><th>Item</th><th>Value</th></tr>')
			for(var o in oElement){
				//if(bHas && !oElement.hasOwnProperty(o)) continue // not included in DOM object
				sHtml.append("\n<tr><td> " + o + " </td>")
				if(oElement[o] == null || oElement[o] == "") sHtml.append('<td>&nbsp;</td></tr>')
				else if(typeof(oElement[o]) == "string") sHtml.append('<td>' + oElement[o] + '</td></tr>')
				else if(typeof(oElement[o]) == "undefined") sHtml.append('<td>undefined</td></tr>')
				else if(__H.isElement(oElement[o])){
					sHtml.append('<td><a href="javascript:void(0)" onclick="__HWindow.setInnerHTML(document.all[' + oElement[o].sourceIndex + '])">Element: ' + oElement[o] + '</a></td></tr>')
				}
				else if(__H.isElementArray(oElement[o])){
					sHtml.append('<td><a href="javascript:void(0)" onclick="__HWindow.setInnerHTML(__H.$N(\'' + oElement[o].length + '\'))">ElementArray: ' + oElement[o] + '</a></td></tr>')
				}
				else if(typeof(oElement[o]) == "object"){
					sHtml.append('<td><a href="javascript:void(0)" onclick="__HWindow.setInnerHTML(' + o + ')">ElementUnknown: ' + oElement[o] + '</a></td></tr>')
				}
				else sHtml.append("<td>" + oElement[o] + " </td></tr>")
			}
			sHtml.append('\n<tr><th>Item</th><th>Value</th></tr>')
			sHtml.append("\n</table>")			
			return sHtml.toString()
		}
		catch(ee){
			__HLog.error(ee,this)
			return "[runtime error]";
		}
	}
})

var __HWindow = new __H.UI.Window()