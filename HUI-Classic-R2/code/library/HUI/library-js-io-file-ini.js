// nOsliw Solutions HUI, HTML Application Library http://hui.nosliw.info, License http://hui.codeplex.com/license
//**Start Encode**

/*

File:     library-js-io-file-ini.js
Purpose:  Development script
Author:   Woody Wilson
Created:  2005-06
Version:  
Dependency:  library-js.js, library-js-io-file.js

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


/**
    Author:     Woody Wilson
	Function:   io_ini_isfile()
	Parameters: sFile,sExt
	            sFile (string) = absolute path name of the file name
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Check if the file has an INF or INF extension
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_isfile(sFile,sExt){
	try{
		sExt = sExt ? "|" + sExt : "";
		var oRe = new RegExp("ini|inf|reg|url" + sExt,"ig")
		if(oFso.FileExists(sFile) && (sExt = oFso.GetExtensionName(sFile))){
			return sExt.search(oRe) > -1
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_parameterexists()
	Parameters: sFile,sParameter,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to find
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Get the line number of a parameter if exists
	Returns:    Boolean (false) or Number
	Exception:  On Error
*/
function io_ini_parameterexists(sFile,sParameter,sExt){
	try{
		var sResult = false
		if(io_ini_isfile(sFile,sExt)){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				if(sParameter == io_ini_isparameter(oFile.ReadLine())){
					sResult = oFile.Line;
					break;
				}
			}
			oFile.Close();
		}
		return sResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
 
/**
    Author:     Woody Wilson
	Function:   io_ini_isparameter()
	Parameters: sParameter
	            sParameter (string) = The parameter or section name to check
	Purpose:    Matches a parameter name in a line. Check if it is a parameter.
	Returns:    Boolean (false) or String
	Exception:  On Error
*/
function io_ini_isparameter(sParameter){
	try{
		if(js_str_blankline(sParameter)) return false;
		var oRe = new RegExp("[\[]([a-z0-9_ \-\.\(\)]+)[\]]$","ig"); // Matches: [paraME TEr_10.A] or [paraME TEr_10.(A)]	
		sParameter = js_str_trim(sParameter)
		if(sParameter.substring(0,1) == ";");
		else if(oRe.exec(sParameter)){	
			return RegExp.$1;
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_keyexists()
	Parameters: sFile,sParameter,sKeyName
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to find
	            sKeyName (string) = Key name to find in parameter
	Purpose:    Gets a key name in a parameter if exists and returns the line number
	Returns:    Boolean (false) or Number
	Exception:  On Error
*/
function io_ini_keyexists(sFile,sParameter,sKeyName){
	try{
		var sResult = false, iLine
		if(iLine = io_ini_parameterexists(sFile,sParameter)){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream && oFile.Line < iLine) oFile.SkipLine();
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine(), oKey;
				if(io_ini_isparameter(sLine)) break; // Loop until next parameter
				else if((oKey = io_ini_iskey(sLine,oFile.Line)) && oKey.name == sKeyName){
					sResult = (oFile.Line-1);
					break;
				}
			}	
			oFile.Close();
		}
		return sResult
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_iskey()
	Parameters: sLine,iLine
	            sLine (string) = The line to check
	            iLine (number,optional) = The line number
	Purpose:    Matches a key line in a parameter. Returns a object contaning name, value etc.
	Returns:    Boolean (false) or Object
	Exception:  On Error
*/
function io_ini_iskey(sLine,iLine){
	try{
		if(js_str_blankline(sLine)) return false;
		var sLineOrg = sLine, sLine = js_str_trim(sLine)
		var oRe = new RegExp("[ \t]+","g");
		sLine = sLine.replace(oRe," "); // Removes all tabs and double spaces
		oRe = new RegExp("[ ]{0,1}=[ ]{0,1}","g"), sLine = sLine.replace(oRe,"="); // Removes spaces near '='
		oRe = new RegExp("[\"'%]{0,1}([^=\"\'%]+)[\"'%]{0,1}={0,1}(.*)","ig"); // This isSearches: anything=any thing  OR  "any thing"=any thing  OR  anything=any thing  OR  anything
		if(sLine.substring(0,1) == ";");
		else if(oRe.exec(sLine)){
			return {
				sline : sLineOrg,
				name  : RegExp.$1,
				value : RegExp.$2 ? RegExp.$2 : "",
				key   : RegExp.$1 + "=" + RegExp.$2,
				iline : iLine ? iLine : false
			}
		}
		return false;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_addparameter()
	Parameters: sFile,sParameter,sKeyName,sKeyValue,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to add
	            sKeyName (string,optional) = Key name to add in parameter
	            sKeyValue (string,optional) = Key value to add in parameter
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Adds a parameter at the end of the file
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_addparameter(sFile,sParameter,sKeyName,sKeyValue,sExt){
	try{
		if(!io_ini_parameterexists(sFile,sParameter,sExt)){
			var oFile = oFso.OpenTextFile(sFile,oReg.append,true,oReg.TristateUseDefault);
			oFile.WriteBlankLines(1);
			oFile.WriteLine("[" + sParameter + "]");
			if(sKeyName && sKeyValue) oFile.WriteLine(sKeyName + " = " + sKeyValue);
			oFile.Close();
			return true;
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_getparameters()
	Parameters: sFile,sExt
	            sFile (string) = absolute path name of the file name
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Get a list of parameters in a file. Returns as an array.
	Returns:    Boolean (false) Or Object (array)
	Exception:  On Error
*/
function io_ini_getparameters(sFile,sExt){
	try{
		if(io_ini_isfile(sFile,sExt)){
			var aParams = [], sLine;
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);					
			while(!oFile.AtEndOfStream){
				if((sLine = oFile.ReadLine()) && (sParameter = io_ini_isparameter(sLine))){
					aParams.push(sParameter);
				}
			}
			oFile.Close();
			return aParams
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_getkeys()
	Parameters: sFile,sParameter,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to find
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Get keys in a parameter of a file. Returns an array key objects containing name, value etc
	Returns:    Boolean (false) Or Object (array) Object
	Exception:  On Error
*/
function io_ini_getkeys(sFile,sParameter,sExt,sKeyOnly){
	var sResult = false
	if(js_str_blankline(sParameter)) return false;
	else if(io_ini_isfile(sFile,sExt) && typeof(sParameter) == "string"){
		try{
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				var p = io_ini_isparameter(sLine);
				if(p && p.toLowerCase() == sParameter.toLowerCase()){
					sResult = [];
					while(!oFile.AtEndOfStream){
						sLine = oFile.ReadLine();							
						if(!io_ini_isparameter(sLine)){ // Get all keys to next parameter
							if(oLine = io_ini_iskey(sLine,oFile.Line)){
								if(sKeyOnly && oLine.name == sKeyOnly){
									return [{
										name  : oLine.name,
										value : oLine.value
									}]
								}
								else sResult.push(oLine);
							}
						}
						else break;
					}
					break;
				}
			}
			
		}
		catch(ee){
			js_log_error(2,e);
			return false;
		}
		finally{
			oFile.Close();
		}
	}	
	return sResult;
}

function io_ini_getkeys_dict(sFile,sParameter,sExt,bIgnoreEmptyValue){
	try{
		var aKeys, oDict = new ActiveXObject("Scripting.Dictionary")
		if(aKeys = io_ini_getkeys(sFile,sParameter,sExt)){
			for(var i = 0, len = aKeys.length; i < len; i++){
				if(!oDict.Exists(aKeys[i].name)){
					if(bIgnoreEmptyValue && (aKeys[i].value).length < 1) continue
					oDict.Add(aKeys[i].name,aKeys[i].value)
				}
			}
		}
		return oDict
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		js_str_kill(aKeys)
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_delparameter()
	Parameters: sFile,sParameter,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to delete
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Delete a parameter in a file if exists
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_delparameter(sFile,sParameter,sExt){
	try{
		var iLine
		if(iLine = io_ini_parameterexists(sFile,sParameter,sExt)){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			var oTmpFile = oFso.OpenTextFile(sFile + ".tmp",oReg.write,true);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				if(oFile.Line < iLine) oTmpFile.WriteLine(sLine);
				else break;
			}
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine()
				if(io_ini_isparameter(sLine)){
					oTmpFile.WriteLine(sLine);
					break;
				}
			}
			while(!oFile.AtEndOfStream) oTmpFile.WriteLine(oFile.ReadLine());					
			oFile.Close();
			oTmpFile.Close();
			io_file_overdelete(sFile + ".tmp",sFile);
			return true;
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_delkey()
	Parameters: sFile,sParameter,sKeyName,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to edit
	            sKeyName (string) = Key name to delete in parameter
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Delete a parameter in a file if exists
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_delkey(sFile,sParameter,sKeyName,sExt){
	try{
		var iLine
		if(iLine = io_ini_keyexists(sFile,sParameter,sKeyName,sExt)){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			var oTmpFile = oFso.OpenTextFile(sFile + ".tmp",oReg.write,true,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				if(oFile.Line != iLine) oTmpFile.WriteLine(sLine);
				else {
					oTmpFile.WriteLine(sLine);
					oFile.SkipLine(); // Delete key
				}
			}							
			oFile.Close();
			oTmpFile.Close();
			io_file_overdelete(sFile + ".tmp",sFile);
			return true;
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_setkey()
	Parameters: sFile,sParameter,sKeyName,sKeyValue,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to edit
	            sKeyName (string) = Key name to set in parameter
	            sKeyValue (string) = Key value to set in parameter
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Set a key name in a parameter of a file
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_setkey(sFile,sParameter,sKeyName,sKeyValue,sExt){
	try{
		var iLine
		if(iLine = io_ini_keyexists(sFile,sParameter,sKeyName,sExt)){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			var oTmpFile = oFso.OpenTextFile(sFile + ".tmp",oReg.write,true,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				if(oFile.Line != iLine) oTmpFile.WriteLine(sLine);
				else {
					oTmpFile.WriteLine(sLine);
					oFile.SkipLine(); // Delete key
					oTmpFile.WriteLine(sKeyName + "=" + sKeyValue); // Change key value
				}
			}							
			oFile.Close();
			oTmpFile.Close();
			io_file_overdelete(sFile + ".tmp",sFile);
			return true;
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_addkey()
	Parameters: sFile,sParameter,sKeyName,sKeyValue,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to edit
	            sKeyName (string) = Key name to set in parameter
	            sKeyValue (string) = Key value to set in parameter
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Add a key name in a parameter of a file
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_addkey(sFile,sParameter,sKeyName,sKeyValue,sExt,bIgnoreExist){
	try{
		sKeyValue = (typeof(sKeyValue) == "string") ? sKeyValue : ""
		if((iLine = io_ini_parameterexists(sFile,sParameter,sExt)) && sKeyName){
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			var oTmpFile = oFso.OpenTextFile(sFile + ".tmp",oReg.write,true,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();
				if(oFile.Line < iLine) oTmpFile.WriteLine(sLine);
				else{
					oTmpFile.WriteLine(sLine); // write the parameter
					break;
				}
			}
			var v = (sKeyValue === "")  ? sKeyName : sKeyName + "=" + sKeyValue
			var oRe = new RegExp("^$","g"); // Removes empty lines
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine();				
				if(!bIgnoreExist && (oKey = io_ini_iskey(sLine,oFile.Line)) && oKey.name == sKeyName){
					oTmpFile.WriteLine(v); // If exist, change key value
					break;
				}
				else if(io_ini_isparameter(sLine) || oFile.AtEndOfStream){ // Loops until next paramter
					if(oFile.AtEndOfStream){
						oTmpFile.WriteLine(sLine); // write the key
						oTmpFile.WriteLine(v); // Add key value at end
						oTmpFile.WriteBlankLines(1);								
					}
					else {
						oTmpFile.WriteLine(v); // Add key value at end
						oTmpFile.WriteBlankLines(1);
						oTmpFile.WriteLine(sLine); // write the next parameter
					}
					break;
				}
				else if(!oRe.exec(sLine)) oTmpFile.WriteLine(sLine);
			}
			while(!oFile.AtEndOfStream){
				oTmpFile.WriteLine(oFile.ReadLine());
			}
			oFile.Close();
			oTmpFile.Close();
			io_file_overdelete(sFile + ".tmp",sFile);
			return true;
		}
		else return io_ini_addparameter(sFile,sParameter,sKeyName,sKeyValue,sExt); // Adds parameter if not exist
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_getkeyvalue()
	Parameters: sFile,sParameter,sKeyName,sExt
	            sFile (string) = The absolute path name of the file name
	            sParameter (string) = The parameter or section name to edit
	            sKeyName (string) = Key name to get in parameter
	            sExt (string,optional) = Extra extension. Could be in Regular expression.
	Purpose:    Get a key value in a parameter of a file
	Returns:    Boolean
	Exception:  On Error
*/
function io_ini_getkeyvalue(sFile,sParameter,sKeyName,sExt){
	try{
		var aKeys
		if(aKeys = io_ini_getkeys(sFile,sParameter,sExt,sKeyName)){
			return (aKeys[0].value).replace(/^"|"$/g,"")
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

/**
    Author:     Woody Wilson
	Function:   io_ini_getkeylines()
	Parameters: sFile
	            sFile (string) = The absolute path name of the file name
	Purpose:    Get all paramter lines. Skip commented lines
	Returns:    Boolean (false) Or Object (array)
	Exception:  On Error
*/
function io_ini_getkeylines(sFile){
	try{
		if(oFso.FileExists(sFile)){
			var a = [];
			var oFile = oFso.OpenTextFile(sFile,oReg.read,false,oReg.TristateUseDefault);
			while(!oFile.AtEndOfStream){
				var sLine = oFile.ReadLine()
				if(typeof(sLine) == "string" && sLine.substring(0,1) != ";" && sLine.substring(0,1) != "#"){
					if(oLine = io_ini_iskey(sLine,oFile.Line)){
						a.push(oLine);
					}
				}
			}
			oFile.Close();
			var a
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_file_ini(iOpt,sFile,sParameter,sKeyName,sKeyValue,sLine,iLine,sExt){
	try{
		var RETURN = false;
		switch(iOpt){
			case 0 : { // Description: Checks validation | Arguments: 0,sFile | Returns: boolean
				RETURN = io_ini_isfile(sFile,sExt)
				break;
			}
			case 1 : { // Description: Checks if parameter exist | Arguments: 1,sFile,sParameter | Returns: line number or false
				RETURN = io_ini_parameterexists(sFile,sParameter,sExt)
				break;
			}
			case 2 : { // Description: Checks if key exist in parameter | Arguments: 2,sFile,sParameter,sItemName,null,null,iLine | Returns: line number or false
				RETURN = io_ini_keyexists(sFile,sParameter,sKeyName)
				break;
			}
			case 3 : { // Description: Appends parameter to file | Arguments: 3,sFile,sParameter,sKeyName,sKeyalue | Returns: Boolean
				RETURN = io_ini_addparameter(sFile,sParameter,sKeyName,sKeyValue)
				break;
			}
			case 4 : { // Description: Gets parameters | Arguments: 4,sFile | Returns: Array of parameters
				RETURN = io_ini_getparameters(sFile,sExt)
				break;
			}
			case 5 : { // Description: Gets keys in parameter | Arguments: 5,sFile,sParameter,null,null,null,iLine | Returns: Array object of key items
				RETURN = io_ini_getkeys(sFile,sParameter)
				break;
			}
			case 6 : { // Description: Deletes a parameter including all its keys | Arguments: 6,sFile,sParameter | Returns:  Boolean
				RETURN = io_ini_delparameter(sFile,sParameter,sExt)
				break;
			}
			case 7 : case 8 : { // Description: Deletes key(iOpt 7) or changes(iOpt 8) a value in a parameter | Arguments: 7/8,sFile,sParameter,sKeyName,sKeyValue | Returns: Boolean
				if(iOpt == 8) RETURN = io_ini_setkey(sFile,sParameter,sKeyName,sKeyValue,sExt)
				else RETURN = io_ini_delkey(sFile,sParameter,sKeyName,sExt)
				break;
			}
			case 9 : { // Description: Adds a key in parameter | Arguments: 9,sFile,sParameter,sKeyName,sKeyValue | Returns: Boolean
				RETURN = io_ini_addkey(sFile,sParameter,sKeyName,sKeyValue,sExt)
				break;
			}
			case 10 : { // Get Parameter
				
				break;
			}
			case 11 : { // Checks string, if is a parameter | Arguments: 11,null,sParameter | Returns: Parameter or false
				RETURN = io_ini_isparameter(sParameter)
				break;
			}
			case 12 : { // Checks string, if is a key | Arguments: 12,null,null,null,null,sLine,iLine | Returns: Key object or false
				RETURN = io_ini_iskey(sLine,iLine)
				break;
			}
			case 13 : { // Gets key value | Arguments: 13,sFile,sParameter,sKeyName | Returns: Key name or false
				RETURN = io_ini_getkeyvalue(sFile,sParameter,sKeyName)
				break;
			}
			case 14 : { // Gets list of lines (keys). | Arguments: 14,sFile | Returns: Array of line
				RETURN = io_ini_getkeylines(sFile)
				break;
			}
			default : {
				break;
			}
		}
		return RETURN;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}
