// nOsliw Solutions HUI, HTML Application Library https://github.com/woowil/HTAFrameworks
//**Start Encode**

/*

File:     library-js-class-file-compression.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2006-04
Version:
Dependency:  library-js.js, library-js__prototype.js

Description:
JScript library scripting functions used for WSH files or HTA Applications.

Usage:

Description:
JScript library scripting functions used for WSH files or HTA Applications.

Revisions: to many

Disclaimer:
This sample code is provided AS IS WITHOUT WARRANTY
OF ANY KIND AND IS PROVIDED WITHOUT ANY IMPLIED WARRANTY
OF FITNESS FOR PURPOSES OF MERCHANTABILITY. Use this code
is to be undertaken entirely at your risk, and the
results that may be obtained from it are dependent on the user.
Please note to fully back up files and system(s) on a regular
basis. A failure to do so can result in loss of data or damage
to systems.

*/


// Global Scripting Objects
try{
	if(!oFso || !oWsh || !oWno){
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oWsh = new ActiveXObject("WScript.Shell");
		var oWno = new ActiveXObject("WScript.Network");
	}
}
catch(ee){
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var oWsh = new ActiveXObject("WScript.Shell");
	var oWno = new ActiveXObject("WScript.Network");
}


// Function _Class() must be loaded
class_file_compression.prototype = new _Class("Compression","Class for compressing javascript/jscript files")

function class_file_compression(sStatus,bDebug){
	try{
		this.isStatus = sStatus ? true : false
		this.isDebug = bDebug ? true : false
		this.encryption = null
		this.delcomments = false
		this.delcomments_all = false
		this.dellines = false
		this.addsemicolon = false
		this.mergelines = false
		this.delwhitespace = false
		this.isRecursive = false
		this.isEncrypt = false

		this.isScript = function(sFile){
			if(typeof(sFile) != "string") return false
			var s
			if(oFso.FileExists(sFile) && (s = oFso.GetExtensionName(sFile))){
				return s.match(/js|jse/ig)
			}
			return false
		}

		this.getFiles = function(sFolder){
			try{
				var sFile = this.fso.temp + "\\class_file_compression_getFiles" + this.random(100,1) + ".log"
				if(oFso.FolderExists(sFolder)){
					if(this.isDebug) this.popup("%comspec% /c dir \"" + sFolder + "\\*.js*\" /b /ON " + (this.isRecursive ? "/s" : "") + ">" + sFile)
					if(oWsh.Run("%comspec% /c dir \"" + sFolder + "\\*.js*\" /b /ON " + (this.isRecursive ? "/s" : "") + ">" + sFile,this.fso.hide,true) == 0){
						var oFile = oFso.OpenTextFile(sFile,this.fso.read,false,this.fso.TristateUseDefault)
						if(this.isDebug) this.popup(oFile.ReadAll())
						var a = (oFile.ReadAll()).split("\r")
						oFile.close()
						for(var i = 0, l = a.length; i < l; i++) a[i] = (a[i]).trim()
						return a
					}
				}
			}
			catch(ee){
				this.error(ee,"getFiles()")
			}
			finally{
				if(oFso.FileExists(sFile)) oFso.DeleteFile(sFile,true)
			}
			return false
		}

		this.compressFolder = function(sSrcFolder,sDstFolder){
			var bResult = true, aFiles
			this.echo("# Using source folder: " + sSrcFolder + " for compress")
			if(!(aFiles = this.getFiles(sSrcFolder))){
				this.echo("## Unable to get jscript files from folder: " + sSrcFolder)
				return false
			}			
			if(!this.isString(sDstFolder)) sDstFolder = sSrcFolder
			for(var i = 0, len = aFiles.length; i < len; i++){
				this.echo("")
				var sFile = (!this.isRecursive ? sSrcFolder + "\\" : "") + aFiles[i]
				if(!this.compressFile(sFile,sDstFolder)) bResult = false
			}
			this.kill(aFiles)
			return bResult
		}

		this.compressFile = function(sFile,sDstFolder){
			if(!this.isScript(sFile)) return false
			this.echo("## Compressing file: " + sFile)
			if(!this.isString(sDstFolder)) sDstFolder = oFso.GetParentFolderName(sFile)
			var oFile = oFso.OpenTextFile(sFile,this.fso.read,false,this.fso.TristateUseDefault)
			var s = oFile.ReadAll()
			oFile.close()
			s = this.delComments(s)
			s = this.delLines(s)
			s = this.addSemiColon(s)
			s = this.delWhiteSpace(s)
			s = this.mergeLines(s)
			if(this.isEncrypt && !this.encryption){
				this.encryption = new class_file_encryption(true)
				s = this.encryption.encryptStream(s)
			}
			return this.setStreamOnFile(sFile,sDstFolder,s)
		}

		this.uncompressFile = function(sFile){

		}

		this.delComments = function(s){
			if(typeof(s) != "string") return ""
			if(!this.delcomments) return s
			this.echo("### Deleting comments fields..")
			s = s.replace(/(?:http|ftp):\x2f\x2f[a-z0-9\-_?\x2f=.]+/ig,"")
			var a = s.split("\r"), s = ""
			var oRe1 = /(.*(?:function |try|if[ \t]*).+[{][ \t]*)\x2f\x2f.+[^"']{0,}[ \t]*[;]{0,}[ \t]*$/g
			var oRe2 = /(.*)[\/]{0,2}.*(?:http|ftp):\x2f\x2f[a-z0-9\-_?\/=.]+(.*)$/ig
			var oRe3 = /.*(?:function|try|if|else|for|while).+(?:http|ftp):\x2f\x2f.*/ig
			var oRe4 = /.*(?:function|try|if|else|for|while).+\x2f\x2f.+["']/g
			var oRe5 = /.*(?:winnt|ldap|file|iis):\x2f\x2f.+/ig
			var oRe6 = /.*["'].*\x2f\x2f.*["'].*/ig
			var oRe7 = /\x2a\x2aStart Encode\x2a\x2a/ig
			for(var i = 0, len = a.length; i < len; i++){
				if((a[i]).match(/\x2f\x2f/g) && !(a[i]).match(oRe7)){
					a[i] = (a[i]).replace(/\x2f\x2f[ \t]*[^\x2f+](["'][;]*)[ \t]*$/,"$1")
					if((a[i]).match(oRe3) || (a[i]).match(oRe4) || (a[i]).match(oRe5) || (a[i]).match(oRe6));
					else if((a[i]).match(oRe1)) a[i] = RegExp.$1
					else if((a[i]).match(/http:|ftp:/ig)) a[i] = (a[i]).replace(oRe2,"$1$2") // removes html
					else if(this.delcomments_all && ((a[i]).match(/.+\x2f\x2f.*["'][ \t]+\+.+$/g) || (a[i]).match(/.+:\x2f\x2f.*["'].*$/g))) {
						if(!(a[i]).match(/[ \t]+\x2f\x2f.*$/g)) a[i] = (a[i]).replace(/[ \t]*\x2f\x2f(.*)$/g,"")
					}
					else if((a[i]).match(/["'].*\x2f\x2f.+["'].*$/g));
					else if((a[i]).match(/(.+[+].*)[ \t]*\x2f\x2f.*(["'].*)$/g,"$1$2"));
					else a[i] = (a[i]).replace(/([^\x2f]*)\s*\x2f\x2f.*$/g,"$1");
					// cleanup of comments
					if(!(a[i]).match(/comspec|cmd|cscript/ig)){
						if(!(a[i]).match(oRe5)) a[i] = (a[i]).replace(/(.+)[ \t]*\x2f\x2f[^+]+(["']*)$/g,"$1$2")
						a[i] = (a[i]).replace(/(.+["'].+)([ \t]*\x2f\x2f.*)(["'][ \t]*[;]*)[ \t]*$/g,"$1$3")
					}
				}
				s = s + a[i] + "\r"
			}
			this.kill(a)

			// Removes '/* .. */'' and '/* ..\r.. */''
			var a = s.split("\r"), s = ""
			for(var i = 0, ii, len = a.length; i < len; i++){
				if((a[i]).match(/\x2f\x2a/g) && !(a[i]).match(oRe7)){
					if((a[i]).match(/\x2f\x2a.*\x2a\x2f/g)){
						if((a[i]).match(/(.*)\x2f\x2a.*\x2a\x2f[ \t]*$/g)) a[i] = RegExp.$1
					}
					else if((ii = (a[i]).search(/\x2f\x2a/g)) != -1 && !(a[i]).match(/["'].*\x2f\x2a.*["'].*/g)){
						s = s + (a[i]).substring(0,ii) + "\r"
						this.echo("### Comments paragragh START at line: " + i + " char: " + ii)
						for(++i; i < len; i++){
							if((ii = (a[i]).search(/\x2a\x2f/g)) != -1){
								s = s + (a[i]).substring(ii+2,(a[i]).length) + "\r"
								this.echo("#### Comments paragragh STOP at line: " + i + " char: " + ii)
								break;
							}
						}
						continue
					}
				}
				s = s + a[i] + "\r"
			}
			this.kill(a)
			
			return s
		}

		this.delLines = function(s){
			if(typeof(s) != "string") return ""
			if(!this.dellines) return s
			this.echo("### Deleting blank lines..")
			s = s.replace(/\t+/g,"")
			s = s.replace(/[ ]+/g," ")
			s = s.replace(/(.*)[\r]*(.*)/g,"$1\r$2")
			return s
		}

		this.addSemiColon = function(s){
			if(typeof(s) != "string") return ""
			if(!this.addsemicolon) return s
			this.echo("### Adding semi-colon..")
			var a = s.split("\r"), s = ""
			var oRe = /(.+[}{])[ \t]*$|(.+\([\w,]+)[ \t]*$/g
			var oRe2 = /(.+)\(.+,[ \t]*$/g
			var oRe7 = /\x2a\x2aStart Encode\x2a\x2a/ig
			for(var i = 0, len = a.length; i < len; i++){
				if(!(a[i]).match(oRe7)){
					if((a[i]).length > 1 && (a[i]).match(oRe));// s = s + RegExp.$1 + "\r"
					else if((a[i]).match(oRe7));
					else if(!(a[i]).match(/.*(?:for|if)[ \t]*\(.+\)[ \t]*/g)) a[i] = (a[i]).replace(/(.+);*[ \t]*$/g,"$1;")
					if(!(a[i]).match(/(?:for|if)[ \t]*\(.*;.*/g) && !(a[i]).match(oRe2)) a[i] = (a[i]).replace(/[;]+([ \t]*[;]+)?/g,";")
					a[i] = (a[i]).replace(/([}{\(][ \t]*);*$/g,"$1")
					if((a[i]).match(/(.+)\([ \t;]*$/g) || (a[i]).match(oRe2)){ // Fixis inner statement functions
						this.echo("#### Inner parameter functions START at line: " + i)
						s = s + a[i] + "\r"
						for(++i; i < len; i++){
							if((a[i]).match(/\)[ \t]*$|\)[ \t]*;.+$/g)){
								this.echo("#### Inner parameter functions STOP at line: " + i)
								if(s.match(/.+["'].*,[ \t]*/g) || s.match(/.+[|]{2}[ \t]*$/g)) s = s + a[i] + ";\r"
								else s = s + ";" + a[i] + ";\r" // don't know why!!
								break;
							}
							else s = s + (a[i]).replace(/[;]+([ \t]*[;]+)?$/g,"") + "\r"
						}
						continue;
					}
					else if((a[i]).match(/(.+[+])[ \t]*;{1,}[ \t]*$/g)) a[i] = RegExp.$1;
				}
				a[i] = (a[i]).replace(/,[ \t]*;[ \t]*$/g,",")
				a[i] = (a[i]).replace(/[|][ \t]*;[ \t]*$/g,"|")
				if((a[i]).match(/[;]{2,}/g) && !(a[i]).match(/for/g)) a[i] = (a[i]).replace(/[;]{2,}/g,";")
				s = s + a[i] + "\r"
			}
			this.kill(a)
			
			return s
		}

		this.delWhiteSpace = function(s){
			if(typeof(s) != "string") return ""
			if(!this.delwhitespace) return s
			this.echo("### Deleting white space..")			
			s = s.replace(/[ \t]*([&=|]{2,3})[ \t]*/g,"$1")
			s = s.replace(/[ \t]*([!+-\x2f]=)[ \t]*/g,"$1")
			s = s.replace(/(.*^[a-z0-9\-_]+)[ \t]*return/g,"$1return");
			s = s.replace(/\)[ \t]*return/g,")return");
			var a = s.split("\r"), s = ""
			var oRe = /comspec.+|\.match[ \t]*\($|\.replace[ \t]*\($|\.test[ \t]*\($|\.split[ \t]*\($|\.search[ \t]*\($/g
			var oRe2 = /=[ \t]*\/.+\/[igm]{,3}.*|RegExp.+/g
			var oRe3 = /.+(?:comspec|cmd|log_|cscript).+/ig
			for(var i = 0, len = a.length; i < len; i++){				
				if(!(a[i]).match(oRe3) && !(a[i]).match(oRe2) && !(a[i]).match(oRe)){
					a[i] = (a[i]).replace(/[ \t]*([;,<>%!~:=])[ \t]*/g,"$1")
					a[i] = (a[i]).replace(/([^+]) ? ([+])/g,"$1$2");
					a[i] = (a[i]).replace(/([+]) ?([^+])/g,"$1$2");
					if(!(a[i]).match(/["'].*[ \t]+-[ \t]*.*["']/ig)){
						a[i] = (a[i]).replace(/([^-]) ?([-])/g,"$1$2");
						a[i] = (a[i]).replace(/([-]) ?([^-])/g,"$1$2");
					}
				}
				if(!(a[i]).match(oRe) && !(a[i]).match(oRe2)) a[i] = (a[i]).replace(/[ \t]+([}{*&|^!?\)])[ \t]*/g,"$1") // special regular expression charactors
				s = s + this.trim(a[i]) + "\r"
			}
			this.kill(a)
			return s;
		}

		this.mergeLines = function(s){
			if(typeof(s) != "string") return ""
			if(!this.mergelines) return s
			if(!this.delcomments) s = this.delComments(s)
			this.echo("### Merging lines (this may take a while..)")
			var a = s.split("\r"), s = ""
			var oRe7 = /\x2a\x2aStart Encode\x2a\x2a/ig
			for(var i = 0, len = a.length; i < len; i++){
				if((a[i]).length > 0) continue
				/*
				if((a[i]).match(oRe7));
				else if(!(a[i]).match(/.+\x2f\x2f.*["'][ \t]+\+.+$/g)){
					a[i] = (a[i]).replace(/([^\x2f]*)\s*\x2f\x2f.*$/g,"$1")
				}
				if((a[i]).match(/(.*)[\)\(\{\}][ \t]*$/g));
				else if(!(a[i]).match(/(.*);[ \t]*$/g)) a[i] = a[i] + ";"
				a[i] = (a[i]).replace(/(.+})[ \t]*;[ \t]*(.*)/g,"$1;$2")
				*/
				
				if((a[i]).match(oRe7)) s = s + a[i] + "\r"
				else if((a[i]).match(/.+[}{;]$/g)) s = s + (a[i]).replace("\r","")
				else s = s + a[i] + "\r"
			}
			
			this.echo("#### Running line bug fix")
			
			s = s.replace(/;((?:else|if|\}))/g,"$1")
			s = s.replace(/([}\)])((?:for|if|else|with|switch))/g,"$1;$2")
			s = s.replace(/continueelse/g,"continue;else")
			s = s.replace(/;\);/gm,");") // BUG fix for complicated: a ? a : (b ? b : c)
			s = s.replace(/;\)([a-z\-_])/g,");$1") // BUG fix for complicated: a ? a : (b ? b : c)
			
			this.kill(a)
			return s
		}

		this.setStreamOnFile = function(sFile,sDstFolder,s){
			try{
				if(!this.isString(sFile,sDstFolder,s)) return false
				sFile = sDstFolder + "\\" + oFso.GetFileName(sFile)				
				this.echo("### Setting compressed stream buffer to file: " + sFile)
				if(!oFso.FolderExists(sDstFolder)){
					oWsh.Run("%comspec% /c md \"" + sDstFolder + "\"",this.fso.hide,true)
				}
				if(oFso.FileExists(sFile)){
					var oFile = oFso.GetFile(sFile); // deletes also empty files
					oFile.Delete();
				}
				//oFso.CreateTextFile(sFile,true,this.fso.TristateUseDefault);
				var oFile = oFso.OpenTextFile(sFile,this.fso.write,true,this.fso.TristateUseDefault);
				var a = s.split("\r"); // Just remember to separate your stream with a Newline character
				for(var i = 0, len = a.length; i < len; i++){
					if((a[i]).length > 0) oFile.WriteLine(a[i]);
				}
				oFile.Close()
				
				this.kill(a,s)
				return true
			}
			catch(ee){
				this.error(ee,"setStreamOnFile()")
			}
		}

		this.compressInfo = function(){

		}
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

