// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions, All Rights Reserved.

//**Start Encode**

__H.include("HUI@IO@File.js","HUI@IO@File_JSEncryption.js")

__H.register(__H.IO.File,"JSCompression","JavaScript/JScript Compressing",function JSCompression(){
	
	var options = {
		b_delcomments     : false,
		b_delcomments_all : false,
		b_dellines		  : false,
		b_addsemicolon	  : false,
		b_mergelines	  : false,
		b_delwhitespace	  : false,
		b_recursive		  : false,
		b_encrypt		  : false
	}
	
	this.initialize = function initialize(oOptions){
		try{
			options.extend(oOptions,true)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.isScript = function isScript(sFile){
		if(typeof(sFile) != "string") return false
		return oFso.FileExists(sFile) && (oFso.GetExtensionName(sFile)).isSearch(/js|jse/ig)
	}

	this.getFiles = function getFiles(sFolder){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			if(oFso.FolderExists(sFolder)){
				//if(__H.$debug) __HLog.popup("%comspec% /c dir \"" + sFolder + "\\*.js*\" /b /ON " + (options.b_recursive ? "/s" : "") + ">" + sFile)
				if(oWsh.Run("%comspec% /c dir \"" + sFolder + "\\*.js*\" /b /ON " + (options.b_recursive ? "/s" : "") + ">" + sFile,__HIO.hide,true) == 0){
					var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault)
					//if(__H.$debug) __HLog.popup(oFile.ReadAll())
					var a = (oFile.ReadAll()).split("\r")
					oFile.close()
					for(var i = 0, l = a.length; i < l; i++) a[i] = (a[i]).trim()
					return a
				}
			}
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			if(oFso.FileExists(sFile)) oFso.DeleteFile(sFile,true)
		}				
	}

	this.compressFolder = function compressFolder(sSrcFolder,sDstFolder){
		try{
			var bResult = true, aFiles
			__HLog.debug("# Using source folder: " + sSrcFolder + " for compress")
			if(!(aFiles = __HFile.listFiles(sSrcFolder,"js"))) return false
			if(!__H.isString(sDstFolder)) sDstFolder = sSrcFolder
			for(var i = 0, iLen = aFiles.length; i < iLen; i++){
				__HLog.debug("")
				var sFile = (!options.b_recursive ? sSrcFolder + "\\" : "") + aFiles[i]
				if(!this.compressFile(sFile,sDstFolder)) bResult = false
				delete aFiles[i]
			}
			return bResult
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}


	this.compressFile = function compressFile(sFile,sDstFolder){
		try{
			if(!this.isScript(sFile)) return false
			__HLog.debug("## Compressing file: " + sFile)
			if(__H.isStringEmpty(sDstFolder)) sDstFolder = oFso.GetParentFolderName(sFile)
			s = __HFile.readall(sFile)
			s = this.delComments(s)
			s = this.delLines(s)
			s = this.addSemiColon(s)
			s = this.delWhiteSpace(s)
			s = this.mergeLines(s)
			if(options.b_encrypt){
				s = __HJSEncrypt.encryptStream(s)
			}
			return this.setStreamOnFile(sFile,sDstFolder,s)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.decompressFolder = function decompressFolder(sSrcFolder,sDstFolder){
		try{
			var bResult = true, aFiles
			__HLog.debug("# Using source folder: " + sSrcFolder + " for decompress")
			if(!(aFiles = __HFile.listFiles(sSrcFolder,"js"))) return false
			if(!__H.isString(sDstFolder)) sDstFolder = sSrcFolder
			for(var i = 0, iLen = aFiles.length; i < iLen; i++){
				__HLog.debug("")
				var sFile = (!options.b_recursive ? sSrcFolder + "\\" : "") + aFiles[i]
				if(!this.decompressFile(sFile,sDstFolder)) bResult = false
				delete aFiles[i]
			}
			return bResult
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.decompressFile = function decompressFile(sFile,sDstFolder){
		try{
			if(!this.isScript(sFile)) return false
			if(__H.isStringEmpty(sDstFolder)) sDstFolder = oFso.GetParentFolderName(sFile)
			s = __HFile.readall(sFile)			
			s = this.decompress(s)
			return this.setStreamOnFile(sFile,sDstFolder,s)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.decompress = function decompress(s){
		try{
			__HLog.debug("Decompressing: ")
			var oReFor = /(for[ \t]*\(.+;.+;.+\)[ \t]\{)/g
			var oReObj = /(.+):(.+),/ig
			var oReHTM = /(&quot;|&amp;|&lt;|&gt;|&nbsp;)$/g
			var TABNUM = 0, ss = "", c, s, tmp, k, sLine, bParam
			
			var tabString = function(num){
				var s = ""
				while(num-- > 0) s = s.concat("\t")
				return s
			}
			
			var ignore = function(j,cEnd){
				var jStep = cEnd.length, jStart = j
				while(sLine.substring(j,j+jStep) != cEnd) j++
				return sLine.substring(jStart,j+jStep)
			}
			
			__HLog.debug("@ 1(3)-Splitting..")
			var a = (s.trim()).replace(oReFor,"$1\n").replace(oReObj,"$1 : $2,\n").split(/\n|\r/g)			
			
			__HLog.debug("@ 2(3)-Breaking.. ")
			for(var i = 0, iLen1 = a.length; i < iLen1; i++){
				sLine = (a[i]).replace(/([a-z0-9_]+)([=]{1,3})(["'a-z0-9_]+)/ig,"$1 $2 $3")
				sLine = sLine.replace(/([a-z0-9_]+)=([\{\[])/ig,"$1 = $2")
				sLine = sLine.replace(/([a-z0-9_"\]\)]+)([+=<>!?]{1,2})(["\[a-z0-9_]+)/ig,"$1 $2 $3")
				sLine = sLine.replace(/,([a-z0-9_]+)/ig,", $1")
				__HLog.debug("@@ Processing line: " + i + "(" + iLen1 + "), length: " + sLine.length)
				for(var j = 0, iLen2 = sLine.length; j < iLen2; j++){					
					switch(c = sLine.substring(j,j+1)){
						case '"' : case '\'' : case '/' : { // NOT WORKING 100%!!
								__HLog.debug("@@ Found " + c + " at col " +j)
								if(c == "\"") c = c + ignore(j,"\"")
								else if(c == "'") c = c + ignore(j,"'")
								else if(c == "/"){
									if(sLine.substring(j,j+2) == "/*") c = ignore(j,"*/")
									else c = ignore(j,"/")
								}
								ss = ss.concat(c)
								j += (c.length-1)
								break;
							}							
						case '{' : {
								__HLog.debug("@@ Found " + c + " at col " +j)
								bParam = true, k = j-1
								if(sLine.substring(j+1,j+2) == "}") bParam = false
								
								while(bParam && k >= 0 && (tmp = sLine.substring(k,k+1))){
									if(tmp.isSearch(/[ \t]/g));
									else if(tmp.isSearch(/[=]/g) && (sLine.substring(k-5,k)).isSearch(/[a-z09_]+[ \t]*$/ig)) break;
									else if(tmp.isSearch(/[\)]/g)) break;
									else if(tmp.isSearch(/[y]/g) && sLine.substring(k-2,k+1) == "try") break;
									else {
										bParam = false
										break;
									}
									k--
								}
								if(bParam){
									ss = ss.concat(c + "\n" + tabString(++TABNUM));
									break;
								}
							}
						case '}' : {
								__HLog.debug("@@ Found " + c + " at col " +j)
								if(TABNUM.isOdd()){
									ss = ss.concat("\n" + tabString(--TABNUM) + c + "\n" + tabString(TABNUM));
									if(sLine.substring(j+1,j+2) == ";") j++
									break;
								}
							}
						case ';' : { 
								__HLog.debug("@@ Found " + c + " at col " +j)
								if(j-8 > 0 && (tmp = sLine.substring(j-8,j)) && tmp.isSearch(oReHTM));
								else {
									ss = ss.concat("\n" + tabString(TABNUM));
									break;
								}								
							}
							
						default  : ss = ss.concat(c); break;
					}
				}
				delete a[i]
			}
			
			__HLog.debug("@ 3(3)-Fixing..")
			var b = ss.split(/\n|\r/g), ss = ""
			var oRe = /[ \t]*(.+)[ \t]*$/g
			for(var i = 0, iLen1 = b.length; i < iLen1; i++){
				if((b[i]).isSearch(/[ \t]*for\(.+/g) && !(b[i]).isSearch(/for\(.+ in .+/g)){
					var s = b[i++] + "; " + (b[i++]).replace(oRe,"$1; ") + (b[i]).replace(oRe,"$1\n")
					ss = ss.concat(s)
					continue
				}
				ss = ss.concat(b[i] + "\n")
				delete b[i]
			}
			return ss
		}
		catch(ee){
			__HLog.error(ee,this)
			return ss
		}
	}


	this.delComments = function delComments(s){
		try{
			if(typeof(s) != "string") return ""
			if(!options.b_delcomments) return s
			__HLog.debug("### Deleting comments fields..")
			s = s.replace(/(?:http|ftp):\x2f\x2f[a-z0-9\-_?\x2f=.]+/ig,"")
			var a = s.split("\r"), s = ""
			var oRe1 = /(.*(?:function |try|if[ \t]*).+[{][ \t]*)\x2f\x2f.+[^"']{0,}[ \t]*[;]{0,}[ \t]*$/g
			var oRe2 = /(.*)[\/]{0,2}.*(?:http|ftp):\x2f\x2f[a-z0-9\-_?\/=.]+(.*)$/ig
			var oRe3 = /.*(?:function|try|if|else|for|while).+(?:http|ftp):\x2f\x2f.*/ig
			var oRe4 = /.*(?:function|try|if|else|for|while).+\x2f\x2f.+["']/g
			var oRe5 = /.*(?:winnt|ldap|file|iis):\x2f\x2f.+/ig
			var oRe6 = /.*["'].*\x2f\x2f.*["'].*/ig
			var oRe7 = /\x2a\x2aStart Encode\x2a\x2a/ig
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if((a[i]).isSearch(/\x2f\x2f/g) && !(a[i]).isSearch(oRe7)){
					a[i] = (a[i]).replace(/\x2f\x2f[ \t]*[^\x2f+](["'][;]*)[ \t]*$/,"$1")
					if((a[i]).isSearch(oRe3) || (a[i]).isSearch(oRe4) || (a[i]).isSearch(oRe5) || (a[i]).isSearch(oRe6));
					else if((a[i]).match(oRe1)) a[i] = RegExp.$1
					else if((a[i]).isSearch(/http:|ftp:/ig)) a[i] = (a[i]).replace(oRe2,"$1$2") // removes html
					else if(options.b_delcomments_all && ((a[i]).isSearch(/.+\x2f\x2f.*["'][ \t]+\+.+$/g) || (a[i]).isSearch(/.+:\x2f\x2f.*["'].*$/g))) {
						if(!(a[i]).isSearch(/[ \t]+\x2f\x2f.*$/g)) a[i] = (a[i]).replace(/[ \t]*\x2f\x2f(.*)$/g,"")
					}
					else if((a[i]).isSearch(/["'].*\x2f\x2f.+["'].*$/g));
					else if((a[i]).isSearch(/(.+[+].*)[ \t]*\x2f\x2f.*(["'].*)$/g,"$1$2"));
					else a[i] = (a[i]).replace(/([^\x2f]*)\s*\x2f\x2f.*$/g,"$1");
					// cleanup of comments
					if(!(a[i]).isSearch(/comspec|cmd|cscript/ig)){
						if(!(a[i]).isSearch(oRe5)) a[i] = (a[i]).replace(/(.+)[ \t]*\x2f\x2f[^+]+(["']*)$/g,"$1$2")
						a[i] = (a[i]).replace(/(.+["'].+)([ \t]*\x2f\x2f.*)(["'][ \t]*[;]*)[ \t]*$/g,"$1$3")
					}
				}
				s = s.concat(a[i] + "\r")
				delete a[i]
			}
			a.length = 0
			// Removes '/* .. */'' and '/* ..\r.. */''
			var a = s.split("\r"), s = ""
			for(var i = 0, ii, iLen = a.length; i < iLen; i++){
				if((a[i]).isSearch(/\x2f\x2a/g) && !(a[i]).isSearch(oRe7)){
					if((a[i]).isSearch(/\x2f\x2a.*\x2a\x2f/g)){
						if((a[i]).match(/(.*)\x2f\x2a.*\x2a\x2f[ \t]*$/g)) a[i] = RegExp.$1
					}
					else if((ii = (a[i]).search(/\x2f\x2a/g)) != -1 && !(a[i]).isSearch(/["'].*\x2f\x2a.*["'].*/g)){
						s = s.concat((a[i]).substring(0,ii) + "\r")
						__HLog.debug("### Comments paragragh START at line: " + i + " char: " + ii)
						for(++i; i < iLen; i++){
							if((ii = (a[i]).search(/\x2a\x2f/g)) != -1){
								s = s.concat((a[i]).substring(ii+2,(a[i]).length) + "\r")
								__HLog.debug("#### Comments paragragh STOP at line: " + i + " char: " + ii)
								break;
							}
						}
						continue
					}
				}
				s = s.concat(a[i] + "\r")
				delete a[i]
			}
			a.length = 0			
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
			return s
		}
	}

	this.delLines = function delLines(s){
		try{
			if(typeof(s) != "string") return ""
			if(!options.b_dellines) return s
			__HLog.debug("### Deleting blank lines..")
			s = s.replace(/\t+/g,"")
			s = s.replace(/[ ]+/g," ")
			s = s.replace(/(.*)[\r]*(.*)/g,"$1\r$2")
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}

	this.addSemiColon = function addSemiColon(s){
		try{
			if(typeof(s) != "string") return ""
			if(!options.b_addsemicolon) return s
			__HLog.debug("### Adding semi-colon..")
			var a = s.split("\r"), s = ""
			var oRe = /(.+[}{])[ \t]*$|(.+\([\w,]+)[ \t]*$/g
			var oRe2 = /(.+)\(.+,[ \t]*$/g
			var oRe7 = /\x2a\x2aStart Encode\x2a\x2a/ig
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if(!(a[i]).isSearch(oRe7)){
					if((a[i]).length > 1 && (a[i]).isSearch(oRe));// s = s + RegExp.$1 + "\r"
					else if((a[i]).isSearch(oRe7));
					else if(!(a[i]).isSearch(/.*(?:for|if)[ \t]*\(.+\)[ \t]*/g)) a[i] = (a[i]).replace(/(.+);*[ \t]*$/g,"$1;")
					if(!(a[i]).isSearch(/(?:for|if)[ \t]*\(.*;.*/g) && !(a[i]).isSearch(oRe2)) a[i] = (a[i]).replace(/[;]+([ \t]*[;]+)?/g,";")
					a[i] = (a[i]).replace(/([}{\(][ \t]*);*$/g,"$1")
					if((a[i]).isSearch(/(.+)\([ \t;]*$/g) || (a[i]).isSearch(oRe2)){ // Fixis inner statement functions
						__HLog.debug("#### Inner parameter functions START at line: " + i)
						s = s.concat(a[i] + "\r")
						for(++i; i < iLen; i++){
							if((a[i]).isSearch(/\)[ \t]*$|\)[ \t]*;.+$/g)){
								__HLog.debug("#### Inner parameter functions STOP at line: " + i)
								if(s.isSearch(/.+["'].*,[ \t]*/g) || s.isSearch(/.+[|]{2}[ \t]*$/g)) s = s.concat(a[i] + ";\r")
								else s = s.concat(";" + a[i] + ";\r") // don't know why!!
								break;
							}
							else s = s.concat((a[i]).replace(/[;]+([ \t]*[;]+)?$/g,"") + "\r")
						}
						continue;
					}
					else if((a[i]).match(/(.+[+])[ \t]*;{1,}[ \t]*$/g)) a[i] = RegExp.$1;
				}
				a[i] = (a[i]).replace(/,[ \t]*;[ \t]*$/g,",").replace(/[|][ \t]*;[ \t]*$/g,"|")
				if((a[i]).isSearch(/[;]{2,}/g) && !(a[i]).isSearch(/for/g)) a[i] = (a[i]).replace(/[;]{2,}/g,";")
				s = s.concat(a[i] + "\r")
				delete a[i]
			}
			a.length = 0
			
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
			return ""
		}
	}

	this.delWhiteSpace = function delWhiteSpace(s){
		try{
			if(typeof(s) != "string") return ""
			if(!options.b_delwhitespace) return s
			__HLog.debug("### Deleting white space..")			
			s = s.replace(/[ \t]*([&=|]{2,3})[ \t]*/g,"$1")
			s = s.replace(/[ \t]*([!+-\x2f]=)[ \t]*/g,"$1")
			s = s.replace(/(.*^[a-z0-9\-_]+)[ \t]*return/g,"$1return");
			s = s.replace(/\)[ \t]*return/g,")return");
			var a = s.split("\r"), s = ""
			var oRe = /comspec.+|\.match[ \t]*\($|\.replace[ \t]*\($|\.test[ \t]*\($|\.split[ \t]*\($|\.search[ \t]*\($/g
			var oRe2 = /=[ \t]*\/.+\/[igm]{,3}.*|RegExp.+/g
			var oRe3 = /.+(?:comspec|cmd|log_|cscript).+/ig
			for(var i = 0, iLen = a.length; i < iLen; i++){				
				if(!(a[i]).isSearch(oRe3) && !(a[i]).isSearch(oRe2) && !(a[i]).isSearch(oRe)){
					a[i] = (a[i]).replace(/[ \t]*([;,<>%!~:=])[ \t]*/g,"$1")
					a[i] = (a[i]).replace(/([^+]) ? ([+])/g,"$1$2").replace(/([+]) ?([^+])/g,"$1$2");
					if(!(a[i]).isSearch(/["'].*[ \t]+-[ \t]*.*["']/ig)){
						a[i] = (a[i]).replace(/([^-]) ?([-])/g,"$1$2").replace(/([-]) ?([^-])/g,"$1$2");
					}
				}
				if(!(a[i]).isSearch(oRe) && !(a[i]).isSearch(oRe2)) a[i] = (a[i]).replace(/[ \t]+([}{*&|^!?\)])[ \t]*/g,"$1") // special regular expression charactors
				s = s.concat((a[i]).trim() + "\r")
				delete a[i]
			}			
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
			return ""
		}
	}

	this.mergeLines = function mergeLines(s){
		try{
			if(typeof(s) != "string") return ""
			if(!options.b_mergelines) return s
			if(!options.b_delcomments) s = this.delComments(s)
			__HLog.debug("Merging lines (this may take a while..)")
			var a = s.split("\r"), s = ""
			var oRe7 = /\x2a\x2aStart Encode\x2a\x2a/ig
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if((a[i]).length > 0) continue
				/*
				if((a[i]).isSearch(oRe7));
				else if(!(a[i]).isSearch(/.+\x2f\x2f.*["'][ \t]+\+.+$/g)){
					a[i] = (a[i]).replace(/([^\x2f]*)\s*\x2f\x2f.*$/g,"$1")
				}
				if((a[i]).isSearch(/(.*)[\)\(\{\}][ \t]*$/g));
				else if(!(a[i]).isSearch(/(.*);[ \t]*$/g)) a[i] = a[i] + ";"
				a[i] = (a[i]).replace(/(.+})[ \t]*;[ \t]*(.*)/g,"$1;$2")
				*/
				
				if((a[i]).isSearch(oRe7)) s = s.concat(a[i] + "\r")
				else if((a[i]).isSearch(/.+[}{;]$/g)) s = s.concat((a[i]).replace("\r",""))
				else s = s.concat(a[i] + "\r")
				delete a[i]
			}
			
			__HLog.debug("@ Running line bug fix")
			
			s = s.replace(/;((?:else|if|\}))/g,"$1").replace(/([}\)])((?:for|if|else|with|switch))/g,"$1;$2")
			s = s.replace(/continueelse/g,"continue;else")
			s = s.replace(/;\);/gm,");") // BUG fix for complicated: a ? a : (b ? b : c)
			s = s.replace(/;\)([a-z\-_])/g,");$1") // BUG fix for complicated: a ? a : (b ? b : c)
						
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
			return ""
		}
	}
	
	this.setStreamOnFile = function setStreamOnFile(sFile,sDstFolder,s){
		try{
			if(!__H.isString(sFile,sDstFolder,s)) return false
			sFile = sDstFolder + "\\" + oFso.GetFileName(sFile)				
			__HLog.debug("Dumping output to file: " + sFile)
			__HFolder.create(sDstFolder)
			if(oFso.FileExists(sFile)){
				var oFile = oFso.GetFile(sFile); // deletes also empty files
				oFile.Delete();
			}
			//oFso.CreateTextFile(sFile,true,__HIO.TristateUseDefault);
			var oFile = oFso.OpenTextFile(sFile,__HIO.write,true,__HIO.TristateUseDefault);
			var a = s.split(/\n|\r/g); // Just remember to separate your stream with a Newline character
			for(var i = 0, iLen = a.length; i < iLen; i++){
				if((a[i]).length > 0) oFile.WriteLine(a[i]);
				delete a[i]
			}
			oFile.Close()
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return true
		}
	}

	this.compressInfo = function compressInfo(){
		try{
			t = 1
		}
		catch(ee){
			__HLog.error(ee,this)
		}
	}	
})

var __HCompression = new __H.IO.File.JSCompression()