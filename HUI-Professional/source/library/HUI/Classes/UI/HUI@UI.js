// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**

__H.register(__H,"UI","UI",function UI(){
	this.b_fullscreen = false
	
	this.winX = window.screenLeft
	this.winY = window.screenTop
	this.winWidth = 0
	this.winHeight = 0
	
	this.screenFull = function(x,yw){
		if(!this.b_fullscreen){
			this.winWidth = document.body.offsetWidth + 8
			this.winHeight = document.body.offsetHeight
			this.winX = window.screenLeft
			this.winY = window.screenTop
			window.moveTo(0,0);
			window.resizeTo(window.screen.width,window.screen.height)	
			this.b_fullscreen = true
		}
		
	}
	
	this.screenRestore = function(){
		if(this.b_fullscreen){
			window.moveTo(this.winX,this.winY);
			window.resizeTo(this.winWidth,this.winHeight);
			this.b_fullscreen = false
		}
	}
	
	this.unselectable = { 
		enable : function(e){
			var e = e ? e : window.event;
	 
			if(e.button != 1){
				if(e.target){
					var targer = e.target;
				}
				else if(e.srcElement){
					var targer = e.srcElement;
				}
	 
				var targetTag = targer.tagName.toLowerCase();
				if((targetTag != "input") && (targetTag != "textarea")){
					return false;
				}
			}
		},
	 
		disable : function () {
			return true;
		}	 
	}

})

var __HUI = new __H.UI()
