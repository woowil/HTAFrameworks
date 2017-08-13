// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.

//**Start Encode**

__H.include("HUI@IO@File.js")

__H.register(__H.IO.File,"JSEncryption","Encrypting JavaScript/JScript/VBscript ",function JSEncryption(){
	var files = new ActiveXObject("Scripting.Dictionary")
	
	this.isScript = function isScript(sFile){
		if(typeof(sFile) != "string") return false
		return oFso.FileExists(sFile) && (oFso.GetExtensionName(sFile)).isSearch(/js|jse/ig)
	}
	
	this.encryptFile = function encryptFile(sFile){
		try{
			if(this.isScript(sFile)){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault)
				var sStream = oFile.ReadAll()
				oFile.close()
				sStream = this.encryptStream(sStream)				
				return this.setStreamOnFile(sFile,sStream)
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}
	
	this.encryptStream = function encryptStream(sStream){
		//TODO: 
		return sStream
	}
	
	this.setStreamOnFile = function setStreamOnFile(sFile,sStream){
		//return this.create(sFile,sStream)
	}	
})

var __HJSEncrypt = new __H.IO.File.JSEncryption()