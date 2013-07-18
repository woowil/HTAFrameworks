// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

// http://www.w3schools.com/ado/default.asp
// http://www.devguru.com/Technologies/ado/quickref/ado_intro.html

/*

File:     library-js-com-ado.js
Purpose:  Development script for ODBC/ADO Connectivity
Author:   Woody Wilson
Created:  2004-08
Version:  see LIB_VERSION

Description:
JScript library scripting functions used for WSH files or HTA Applications.
 
Usage:
Need library-js.js

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

var LIB_NAME    = "Library ADO";
var LIB_VERSION = "1.0";
var LIB_FILE    = oFso.GetAbsolutePathName("library-js-sys-ado.js")

var ado = new ado_object();

function ado_object(){
	try{
		this.version = new js_objectversion(LIB_NAME,LIB_FILE,LIB_VERSION)
	}
	catch(ee){}
}

function ado_db_connect(sOpt,iOpt,sServer,sUsername,sPassword,sDatabaseName,sDatabaseFile,sDSNname,sADClient,iTimeout){
	try{ // http://support.microsoft.com/default.aspx?scid=kb;en-us;Q300382&ID=kb;en-us;Q300382&SD=MSDN
		var sConnect = false;
		var oConnect = new ActiveXObject("ADODB.Connection");
		oConnect.ConnectionTimeout = !isNaN(iTimeout) ? iTimeout : 30;
		
		var oFile = new Object();		
		switch(sOpt = sOpt.toLowerCase()){
			case "access" : { // Microsoft Access
				if(sDatabaseFile) oFile = io_file_info(sDatabaseFile);
				if(iOpt == 1 && oFile.ext == "mdb") sConnect = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" + sDatabaseFile; // Without DSN 
				else if(iOpt == 2 && oFile.ext == "mdb") sConnect = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" + sDatabaseFile; // OLE DB
				else if(iOpt == 3) sConnect = "FILEDSN=ADSN"; // File DSN
				else if(iOpt == 4) sConnect = sDSNname; // With DSN and no User ID/Password
				else if(iOpt == 5) return oConnect.Open(sDSNname,sUsername,sPassword); // With DSN and User ID/Password
				else if(iOpt == 6 && oFile.ext == "mdb") sConnect = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" + sDatabaseFile; // Without DSN, using a physical path as a reference
				else return false;				
				break;
			}
			case "mssql" : { // Microsoft SQL Server
				if(iOpt == 1) sConnect = "PROVIDER=SQLOLEDB;DATA SOURCE=" + sServer + ";UID=" + sUsername + ";PWD=" + sPassword + ";DATABASE=" + sDatabaseName; // OLE DB
				else if(iOpt == 2) sConnect = "DSN=" + sDSNname + ";UID=" + sUsername + ";PWD=" + sPassword + ";DATABASE=" + sDatabaseName; // With DSN
				else if(iOpt == 3) sConnect = "DRIVER={SQL Server};SERVER=" + sServer + ";UID=" + sUsername + ";PWD=" + sPassword + ";DATABASE=" + sDatabaseName; // Without DSN
				else return false;
				break;
			}
			case "msfoxpro" : { // Microsoft Visual FoxPro
				if(iOpt == 1 && (oFile = io_file_info(sDatabaseFile)) && oFile.ext == "dbc"){
					sConnect = "Driver=Microsoft Visual Foxpro Driver; UID=" + sUsername + ";SourceType=DBC;SourceDB=" + sDatabaseName; // Without DSN
				}
				else return false;
				break;
			}
			case "oracle" : { // Oracle
				if(iOpt == 1){ // ODBC with DSN
					oConnect.CursorLocation = sADClient;
					// requires use of adovbs.inc; numeric value is 3
					sConnect = "DSN=" + sDSNname + ";UID=" + sUsername + ";PWD=" + sPassword;
				}
				else if(iOpt == 2){ // OLE DB
					oConnect.CursorLocation = sADClient;
					// requires use of adovbs.inc; numeric value is 3
					sConnect = "Provider=MSDAORA.1;Password=" + sPassword + ";User ID=" + sUsername + ";Data Source=data.world";
				}
				else return false;
				break;
			}
			case "mysql" : {
				
				break;
			}
			case "sybase" : {
			
				break;
			}
			default : break;
		}
		// 
		if(sConnect){
			oConnect.ConnectionString = sConnect;
			oConnect.Open();
			return oConnect; // Remember to close connection when finnish: oConnect.Close()
		}
		else return false;
	}
	catch(e){
		ado_log_error(2,e);
		return false;
	}
}

function ado_tbl_query(oConnect,sQuery,bError,iOpt){
	try{ // http://www.w3schools.com/ado/ado_ref_recordset.asp
		ado.err.qry = "";
		var oRecord = new ActiveXObject("ADODB.Recordset");
		oRecord.Open(sQuery,oConnect); // http://www.w3schools.com/ado/met_rs_open.asp		
		return oRecord; // Always remember to close object		
	}
	catch(e){
		if(bError) ado_err_connect(iOpt,oConnect); // Shows Error if generated
		ado.err.qry = sQuery;
		ado_log_error(2,e);
		return false;
	}
}

function ado_tbl_select(oConnect,sQuery,sDelim,bClose,bError,iOpt,sOpt1,sOpt2,sServer,sUsername,sPassword,sDatabase){
	try{
		// If already connected: ado_tbl_select(oConnect,'select * from Table','\t',true,true,1)
		// iOpt = 1: Prints to StdOut via WScript.Echo
		// iOpt = 2: Returns a HTML Table (has defined stylesheet classes in TD tags)
		// iOpt = 3: Returns an 2D array object where ex. MySelectArray["FirstFieldName"][0] = "first cell in first row" 
		// iOpt = 4: Returns a INI file stream where section name is either [ROW-X] (X=1...N) or [sOpt2]
		oConnect = oConnect ? oConnect : ado_db_connect((sOpt2 ? sOpt2 : "mssql"),1,sServer,sUsername,sPassword,sDatabase);
		var oRecord = false, sRow = sIni = sHtml = "";
		if(oRecord = ado_tbl_query(oConnect,sQuery,bError,iOpt)){
			var i = 0, iLen = oRecord.Fields.Count, r = 1;
			sDelim = sDelim ? sDelim : " ";
			sHtml = '<table border=0 cellpadding=0 cellspacing=0 class="cSelectTbl">';
			if(iOpt == 2) sHtml += '\n<tr class="sSelectHd">';
			else if(iOpt == 3) var aField = new Array(), aRows = new Array();
			for(var oField, oEnum = new Enumerator(oRecord.Fields); !oEnum.atEnd(); oEnum.moveNext()){
				oField = oEnum.item();
				if(i++ < (iLen-1)) sRow += oField.Name + sDelim;
				else sRow += oField.Name;
				if(iOpt == 2) sHtml += '\n\t<td>' + oField.Name + '</td>';
				else if(iOpt == 3){
					aField.push(oField.Name);
					aRows[oField.Name] = new Array();
					// http://www.w3schools.com/ado/prop_type.asp
					aRows[oField.Name].name = oField.Name;
					aRows[oField.Name].status = oField.Status;
					aRows[oField.Name].attr = oField.Attributes;
					aRows[oField.Name].type = oField.Type;
					aRows[oField.Name].defsize = oField.DefinedSize;
				}
			}
			if(iOpt == 1) WScript.Echo(sRow), sRow = ""; // WScript
			else if(iOpt == 2) sHtml += '\n</tr>';
			
			while(!oRecord.EOF){
				if(iOpt == 2) sHtml += '\n<tr class="cSelectTr">';
				else if(iOpt == 4){
					sIni += sOpt1 ? "[" + sOpt1 + "]\n" : "[ROW-" + r++ + "]\n";
				}
				
				for(var i = 0, sRow = "", oField, oEnum = new Enumerator(oRecord.Fields); !oEnum.atEnd(); oEnum.moveNext(), i++){
					oField = oEnum.item();
					sRow += oField.Value + ((i < (iLen-1)) ? sDelim : "");
					if(iOpt == 2) sHtml += '\n\t<td class="cSelectTd">' + oField.Value + '</td>';
					else if(iOpt == 3) aRows[oField.Name].push(oField.Value);
					else if(iOpt == 4) sIni += oField.Name + " = " + oField.Value + "\n";
				}
				
				if(iOpt == 1) WScript.Echo(sRow); // WScript
				else if(iOpt == 2) sHtml += "\n</tr>";
				else if(iOpt == 4) sIni += "\n";
				
				oRecord.MoveNext();
			}
			if(iOpt == 2) sHtml += '\n</table>';
		}
		else return false;
		if(bClose) ado_tbl_close(oConnect,oRecord); // Closes connect
	}
	catch(e){
		if(bError) ado_err_connect(iOpt,oConnect); // Shows Error if generated
		ado_log_error(2,e);
		return false;
	}
	finally{
		if(iOpt == 1) return true;
		else if(iOpt == 2) return sHtml;
		else if(iOpt == 3) return aRows;
		else if(iOpt == 4) return sIni;
		else return false;
	}
}

function ado_tbl_insert(oConnect,sInsert,bClose,bError,iOpt,sOpt,sServer,sUsername,sPassword,sDatabase){
	try{
		oConnect = oConnect ? oConnect : ado_db_connect((sOpt ? sOpt : "mssql"),1,sServer,sUsername,sPassword,sDatabase);
		var oRecord = false;
		if(!(oRecord = ado_tbl_query(oConnect,sInsert,bError,iOpt))){
			return false;
		}
		if(bClose) ado_tbl_close(oConnect,oRecord); // Closes connect
		return true;
	}
	catch(e){
		if(bError) ado_err_connect(iOpt,oConnect); // Shows Error if generated
		ado_log_error(2,e);
		return false;
	}
}

function ado_tbl_count(oConnect,sTable,bClose,bError,iOpt,sOpt,sServer,sUsername,sPassword,sDatabase){
	try{
		oConnect = oConnect ? oConnect : ado_db_connect((sOpt ? sOpt : "mssql"),1,sServer,sUsername,sPassword,sDatabase);
		var oRecord = iCount = false;
		if(!(oRecord = ado_tbl_query(oConnect,"SELECT COUNT(*) FROM " + sTable,bError)),iOpt){
			return false;
		}
		oRecord.MoveFirst();
		iCount = oRecord.Fields(0);
		if(bClose) ado_tbl_close(oConnect,oRecord); // Closes connect
	}
	catch(e){
		if(bError) ado_err_connect(iOpt,oConnect); // Shows Error if generated
		ado_log_error(2,e);
		return false;
	}
	finally{
		return iCount;
	}
}

function ado_tbl_close(sOjectStream){
	try{
		for(var i = 0; i < arguments.length; i++){
			if(arguments[i]){
				try{
					arguments[i].Close();
					arguments[i] = null;
				}
				catch(ee){}
			}
		}
		return true;
	}
	catch(e){
		ado_log_error(2,e);
		return false;
	}
}
		
///////////////
/// ERROR

function ado_err_connect(iOpt,oConnect){
	try{ // http://www.w3schools.com/ado/ado_ref_error.asp
		var sText = "";
		for(var oError, oEnum = new Enumerator(oConnect.Errors); !oEnum.atEnd(); oEnum.moveNext()){
			oError = oEnum.item();
			if(iOpt == 1){
				WScript.Echo("Description:\t" + oError.Description);
				WScript.Echo("Help Context:\t" + oError.HelpContext);
				WScript.Echo("Help File:\t" + oError.HelpFile);
				WScript.Echo("Native Error:\t" + oError.NativeError);
				WScript.Echo("Number:\t" + oError.Number);
				WScript.Echo("Source:\t\t" + oError.Source);
				WScript.Echo("SQL State:\t" + oError.SQLState);
				WScript.Echo("");
			}
			else if(iOpt == 2 || iOpt == 3 || iOpt == 4){
				sStr = (iOpt == 2) ? "<br>" : "\n";
				sText += "<br>Description:\t" + oError.Description + sStr;
				sText += "Help Context:\t" + oError.HelpContext + sStr;
				sText += "Help File:\t" + oError.HelpFile + sStr;
				sText += "Native Error:\t" + oError.NativeError + sStr;
				sText += "Number:\t\t" + oError.Number + sStr;
				sText += "Source:\t\t" + oError.Source + sStr;
				sText += "SQL State:\t" + oError.SQLState + sStr;
				sText += "\n";
			}
			ado.err.description = oError.Description;
			ado.err.helpcontext = oError.HelpContext;
			ado.err.helpfile = oError.HelpFile;
			ado.err.nativeerror = oError.NativeError;
			ado.err.number = oError.Number;
			ado.err.source = oError.Source;
			ado.err.sqlstate = oError.SQLState;
		}
		if(iOpt == 2 || iOpt == 3 || iOpt == 4) return sText;
		else return true;
	}
	catch(e){
		ado_log_error(2,e);
		return false;
	}
}

function ado_log_error(iOpt,oErr){
	try{
		js_log_error(iOpt,oErr);
		try{
			js.err.textarea_error.innerText += "\nQUERY:\t\t" + ado.err.qry;
		}
		catch(ee){}
	}
	catch(e){
		try{
			WScript.Echo(oErr.description);
		}
		catch(e){
			alert(oErr.description);
		}
	}
}

