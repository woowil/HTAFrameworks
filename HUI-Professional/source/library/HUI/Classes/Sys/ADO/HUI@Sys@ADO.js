// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


// http://www.w3schools.com/ado/default.asp
// http://www.devguru.com/Technologies/ado/quickref/ado_intro.html
// http://support.microsoft.com/default.aspx?scid=kb;en-us;Q300382&ID=kb;en-us;Q300382&SD=MSDN

// http://www.connectionstrings.com/

__H.include("HUI@Sys.js")

__H.register(__H.Sys,"ADO","Data Access (ADO)",function ADO(){
	var o_this = this
	var b_initialized = false
	
	var d_connection
	var o_connection_tmp
	
	/////////////////////////////////////
	//// DEFAULT
	
	var o_options = {
		i_timeout : 30,
		s_cur_database : "msaccesss"
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;
		
		d_connection = new ActiveXObject("Scripting.Dictionary")
		
		b_initialized = true
	}
	initialize()
	
	this.setOptions = function setOptions(oOptions){		
		try{
			Object.extend(o_options,oOptions,true)
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	/////////////////////////////////////
	//// 
	
	this.isADO = function isADO(){
		try{
			return ;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.setConnect = function setConnect(oConnect,sConnect,sCursor){
		try{
			oConnect.ConnectionTimeout = o_option.i_timeout
			oConnect.ConnectionString  = sConnect
			if(sCursor) oConnect.CursorLocation = sCursor
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getConnect = function getConnect(sConnect,sCursor){
		try{
			if(!d_connection.Exists(sConnect)){
				var oConnect = new ActiveXObject("ADODB.Connection");
				this.setConnect(oConnect,sConnect,sCursor)
				oConnect.Open()
				d_connection.Add(sConnect,oConnect)
			}
			return (o_connection_tmp = d_connection(sConnect))
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.connect = function connect(){
		try{
			switch(o_options.s_cur_database.toLowerCase()){
				case "access"   : 
				case "msaccess" : return this.connectAccess.apply(this,arguments)
			}
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.connectAccess = function connectAccess(sFile,sUser,sPass){
		try{
			var oRe = /mdb|dbc/ig
			if(!oFso.FileExists(sFile) || !oRe.test(oFso.GetExtensionName(sFile))) return __HExp.ArgumentIllegal("Unable to locate access file " + sFile + " or invalid file extension")
			sUser = typeof(sUser) == "string" ? sUser : "Admin"
			sPass = typeof(sPass) == "string" ? sPass : ""

			var sConnect1 = "DRIVER={Microsoft Access Driver (*.mdb)};Password='" + sPass + "';User ID=" + sUser + ";DBQ=" + sFile.toLowerCase()
			var sConnect2 = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;Password='" + sPass + "';User ID=" + sUser + ";DATA SOURCE=" + sFile.toLowerCase()

			return this.getConnect(sConnect1) || this.getConnect(sConnect2)
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}

	this.connectAccessDSN = function connectAccessDSN(sDSN,sUser,sPass){
		try{
			if(typeof(sUser) == "string") return __HExp.ArgumentIllegal("Argument[0] is not a valid DSN name")
			sUser = typeof(sUser) == "string" ? sUser : ""
			sPass = typeof(sPass) == "string" ? sPass : ""

			var sConnect = "FILEDSN=" + sDSN + ";Password='" + sPass + "';User ID=" + sUser + ";"

			return this.getConnect(sConnect)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.connectMSSQL = function connectMSSQL(sServer,sDatabase,sUser,sPass){
		try{
			if(!this.isString(sServer,sDatabase)) return false
			sUser = typeof(sUser) == "string" ? sUser : "Admin"
			sPass = typeof(sPass) == "string" ? sPass : ""

			var sConnect1 = "DRIVER={SQL Server};SERVER=" + sServer + ";UID=" + sUser + ";PWD=" + sPass + ";DATABASE=" + sDatabase.toLowerCase()
			var sConnect2 = "PROVIDER=SQLOLEDB;DATA SOURCE=" + sServer + ";UID=" + sUser + ";PWD=" + sPass + ";DATABASE=" + sDatabase.toLowerCase()

			return (this.getConnect(sConnect1) || this.getConnect(sConnect2))
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.connectMSSQLDSN = function connectMSSQLDSN(sDSN,sDatabase,sUser,sPass){
		try{
			if(!this.isString(sDSN,sDatabase)) return false

			var sConnect = "DSN=" + sDSN + ";UID=" + sUser + ";PWD=" + sPass + ";DATABASE=" + sDatabase
			return this.getConnect(sConnect)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.connectMSFoxPro = function connectMSFoxPro(sDatabase,sUser,sPass){
		try{
			var sConnect = "Driver=Microsoft Visual Foxpro Driver; UID=" + sUser + ";SourceType=DBC;SourceDB=" + sDatabase
			return this.getConnect(sConnect)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.connectOracle = function connectOracle(sSID,sUser,sPass){
		try{
			var sConnect1 = "Provider=MSDAORA.1;Password=" + sPass + ";User ID=" + sUser + ";Data Source=" + sSID

			return this.getConnect(sConnect1) || this.getConnect(sConnect2)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.connectOracleDSN = function connectOracleDSN(sDSN,sUser,sPass){
		try{
			if(!this.isString(sDSN)) return false
			sUser = typeof(sUser) == "string" ? sUser : ""
			sPass = typeof(sPass) == "string" ? sPass : ""

			var sConnect1 = "DSN=" + sDSN + ";UID=" + sUser + ";PWD=" + sPass
			// Oracle Default passwords:
			// SYS / change_on_install
			// SYSTEM / manager
			// Scott / Tiger
			return this.getConnect(sConnect1)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.connectMySQL = function connectMySQL(sServer,sDB,sUser,sPass){
		try{
			var sConnect1 = "Driver={mySQL};Server=" + sServer + ";Option=131072;Port=3309;Stmt=;Database=" + sDBS + "; User=" + sUser + ";Password=" + sPass;
			return this.getConnect(sConnect1)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	//////////////////////

	var select = function select(oRecord,iOpt){
		try{
			var i = 0, iLen = oRecord.Fields.Count, r = 1;
			var OUTPUT_ECHO = (iOpt == 1)
			var OUTPUT_HTML = (iOpt == 2)
			var OUTPUT_ARRAY = (iOpt == 3)
			var OUTPUT_INI = (iOpt == 4)

			var sRow = ""
			var sHtml = '<table border=0 cellpadding=0 cellspacing=0 class="cSelectTbl">\n<tr class="sSelectHd">'
			var aField = [], aRows = [];
			var sIni = ""
			var sDelim = "::"

			for(var oEnum = new Enumerator(oRecord.Fields); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(i++ < (iLen-1)) sRow = sRow.concat(oItem.Name + sDelim);
				else sRow = sRow.concat(oItem.Name)

				if(OUTPUT_HTML) sHtml = sHtml.concat('\n\t<td>' + oItem.Name + '</td>')
				else if(OUTPUT_ARRAY){
					aField.push(oItem.Name);
					aRows[oItem.Name] = []; // http://www.w3schools.com/ado/prop_type.asp
					aRows[oItem.Name].Name = oItem.Name;
					aRows[oItem.Name].Status = oItem.Status;
					aRows[oItem.Name].Attributes = oItem.Attributes;
					aRows[oItem.Name].Type = oItem.Type;
					aRows[oItem.Name].DefinedSize = oItem.DefinedSize;
				}
			}
			
			if(OUTPUT_ECHO) WScript.Echo(sRow)
			else if(OUTPUT_HTML) sHtml = sHtml.concat("\n</tr>")

			while(!oRecord.EOF){
				if(OUTPUT_HTML) sHtml = sHtml.concat('\n<tr class="cSelectTr">')
				else if(OUTPUT_INI){
					sIni = sIni.concat("[ROW-" + r++ + "]\n")
				}

				for(var i = 0, sRow = "", oEnum = new Enumerator(oRecord.Fields); !oEnum.atEnd(); oEnum.moveNext(), i++){
					var oItem = oEnum.item();
					sRow = sRow.concat(oItem.Value + ((i < (iLen-1)) ? sDelim : ""))
					if(OUTPUT_HTML) sHtml = sHtml.concat('\n\t<td class="cSelectTd">' + oField.Value + '</td>')
					else if(OUTPUT_ARRAY) aRows[oItem.Name].push(oItem.Value);
					else if(OUTPUT_INI) sIni = sIni.concat(oItem.Name + " = " + oItem.Value + "\n")
				}

				if(OUTPUT_ECHO) WScript.Echo(sRow);
				else if(OUTPUT_HTML) sHtml = sHtml.concat("\n</tr>")
				else if(OUTPUT_INI) sIni = sIni.concat("\n")

				oRecord.MoveNext();
			}
			if(OUTPUT_HTML) sHtml = sHtml.concat('\n</table>')
			
			if(OUTPUT_ECHO) return true;
			else if(OUTPUT_HTML) return sHtml;
			else if(OUTPUT_ARRAY) return aRows;
			else if(OUTPUT_INI) return sIni;
			
			return false
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
		}
	}

	this.selectEcho = function selectEcho(sQuery){
		try{
			return select(this.query(sQuery),1)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.selectHTML = function selectHTML(sQuery){
		try{
			return select(this.query(sQuery),2)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.selectArray = function selectArray(sQuery){
		try{
			//Returns an 2D array object where ex. MySelectArray["FirstFieldName"][0] = "first cell in first row"
			return select(this.query(sQuery),3)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.selectINI = function selectINI(){
		try{
			// Returns a INI file stream where section name is either [ROW-X] (X=1...N) or [sOpt2]
			return select(this.query(sQuery),4)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	//////////////////////

	this.query = function query(sQuery){
		try{
			var oRecord = new ActiveXObject("ADODB.Recordset");
			oRecord.Open(sQuery,o_connection_tmp); // http://www.w3schools.com/ado/met_rs_open.asp
			return oRecord; // remember to close object
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.insert = this.query;

	this.count = function count(oConnect){
		try{
			var oRecord, iCount = 0;
			if(!(oRecord = this.query("Select count(*) from " + sTable))){
				return iCount;
			}
			oRecord.MoveFirst();
			return oRecord.Fields(0);
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
		finally{
			this.close(oRecord)
		}
	}

	/////////////////////

	this.close = function close(){
		try{
			for(var i = 0, iLen = arguments.length; i < iLen; i++){
				if(typeof(arguments[i]) == "object"){
					try{
						arguments[i].Close();
						arguments[i] = null;
					}
					catch(e1){}
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getError = function getError(oError){
		try{
			var o = {
				Description : oError.Description,
				HelpContext : oError.HelpContext,
				HelpFile    : oError.HelpFile,
				NativeError : oError.NativeError,
				Number      : oError.Number,
				Source      : oError.Source,
				SQLState    : oError.SQLState
			}
			return o
		}
		catch(ee){
			__HLog.error(ee,this)
			return null
		}
	}
})

var __HADO = new __H.Sys.ADO()