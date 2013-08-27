// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**

__H.include("HUI@UI@Window.js")

__H.register(__H.UI.Window,"Elements","HTML Elements",function Elements(){

	this.isElements = function isElements(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(arguments[i] && typeof(arguments[i]) == "object" && arguments[i].style){
				continue
			}
			return false
		}
		return !!i
	}

	this.center = function center(oElements){
		if(!this.isElements(oElements)) return false
		oElements.style.setExpression("left","document.body.clientWidth/2 - oElements.offsetWidth/2");
		oElements.style.setExpression("top","document.body.clientHeight/2 - oElements.offsetHeight/2");
		document.recalc(true);
	}

	this.isVisible = function isVisible(oElements){
		if(!this.isElements(oElements)) return false
		return (oElements.style.visibility == 'visible' || oElements.style.display == "block")
	}

	this.isHidden = function isHidden(oElements){
		if(!this.isElements(oElements)) return false
		return (oElements.style.visibility == "hidden" || oElements.style.display == "none")
	}

	this.show = function show(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!this.isElements(arguments[i])) continue
			arguments[i].style.display = "block";
			arguments[i].style.visibility = "visible";
		}
	}

	this.showImg = function showImg(oElements,oImg,sSrc,bHideElements){
		if(typeof(oElements) != "object" || typeof(oImg) != "object");
		else if(oElements.style && oImg.tagName == "IMG"){
			if(bHideElements) this.show(oElements)
			oImg.src = sSrc
			return true;
		}
		return false;
	}

	this.hide = function hide(){
		for(var i = 0, iLen = arguments.length; i < iLen; i++){
			if(!this.isElements(arguments[i])) continue
			arguments[i].style.display = "none";
			arguments[i].style.visibility = "hidden";
		}
	}

	this.hideImg = function hideImg(oElements,oImg,sSrc,bHideElements){
		if(this.isElements(oElements,oImg) && oImg.tagName == "IMG"){
			if(bHideElements) this.hide(oElements)
			oImg.src = sSrc
			return true;
		}
		return false;
	}

	this.toggle = function toggle(obj){
		if(this.isElements(obj)){
			if(this.isHidden(obj)) this.show(obj);
			else this.hide(obj);
		}
		else if(typeof(obj) == "object" && obj.length){
			for(var o in obj){
				if(obj.hasOwnProperty(oo)){
					this.toggle(obj[o])
				}
			}
		}
	}

	this.toggleImg = function toggleImg(oElements,oImg,sSrcVisible,sSrcHidden,bHideElements){
		if(this.isElements(oElements,oImg) && oImg.tagName == "IMG"){
			if(this.isVisible(oElements)) return this.hideImg(oElements,oImg,sSrcHidden,bHideElements)
			else return this.showImg(oElements,oImg,sSrcVisible,bHideElements)
		}
		return false;
	}

	this.write = function write(oElements,sHtml){
		if(!this.isElements(oElements)) return false
		if(this.bIE || this.bNS) oElements.innerHTML = sHtml
		else if(this.bNSC){
			oElements.open();
			oElements.writeln(sHtml);
			oElements.close();
		}
	}

	this.transition = function transition(obj){
		if(this.isElements(obj) && obj.filters){
			this.hide(obj);
			obj.filters.item(0).stop();
			obj.filters.item(0).apply();
			//obj.filters.item(0).duration = dur ? dur : '3.0';
			//obj.filters.item(0).transition = (trans ? trans : '23');
			this.show(obj);
			obj.filters(0).play(1.000);
		}
	}

	this.select = function select(obj){
		if(this.isElements(obj) && obj.contentEditable){
			if((obj.nodeName).isSearch(/TEXTAREA|INPUT/ig)) obj.select();
			obj.document.execCommand('copy');
		}
	}
	
	this.purge = function purge(oElements){
		// Handlers memory leak in IE
		// http://javascript.crockford.com/memory/leak.html__H
		if(!this.isElements(oElements)) return false
	    var a = oElements.attributes, l;
		if(a){
			l = a.length;
			for(var i = 0; i < l; i += 1) {
				var n = a[i].name;
	            if(typeof oElements[n] === 'function'){
					oElements[n] = null;
				}
			}
		}
	    a = oElements.childNodes;
	    if(a) {
			l = a.length;
			for(var i = 0; i < l; i += 1) {
				this.purge(oElements.childNodes[i]);
			}    
		}
	}
})


//// Global define
var __HElem = new __H.UI.Window.Elements()