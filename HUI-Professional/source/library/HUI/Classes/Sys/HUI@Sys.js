// nOsliw HUI - HTML/HTA Application Framework Library (http://hui.codeplex.com/)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.
// License: GNU Library General Public License (LGPL) (http://hui.codeplex.com/license)
//**Start Encode**


__H.register(__H,"Sys","System",function Sys(){
	var o_this = this;
	var b_initialized = false
	
	var d_ipconfig = null
	
	this.resource	= null
	this.computer	= null
	this.userdomain = null
	this.username	= null
	this.domainuser = null
	this.password	= null
	this.isadmin = true
	
	var hkcr = "HKEY_CURRENT_ROOT"
	var hkcu = "HKEY_CURRENT_USER"
	var hklm = "HKEY_LOCAL_MACHINE"
	var hku = "HKEY_USERS"
	var hkcc = "HKEY_CURRENT_CONFIG"
	
	var a_computernames = [
		"HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\ComputerName\\ComputerName", // Computername
		"HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\ComputerName\\ActiveComputerName",
		"HKEY_LOCAL_MACHINE\\CLONE\\CLONE\\Services\\Tcpip\\Parameters", // Key: Hostname
		"HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\Tcpip\\Parameters", // Key: Hostname
		"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate\\OemInfo"
	]
	
	this.reg_keys = {
		HKCR : 0x80000000,
		HKCU : 0x80000001,
		HKLM : 0x80000002,
		HKU : 0x80000003,
		HKCC : 0x80000005,

		REG_SZ : 1,
		REG_EXPAND_SZ : 2,
		REG_BINARY : 3,
		REG_DWORD : 4,
		REG_DWORD_BIG_ENDIAN : 5,
		REG_LINK : 6,
		REG_MULTI_SZ : 7,
		REG_RESOURCE_LIST : 8,
		REG_FULL_RESOURCE_DESCRIPTOR : 9,
		REG_RESOURCE_REQUIREMENTS_LIST : 10,
		REG_QWORD : 11,

		KEY_QUERY_VALUE : 0x1,
		KEY_SET_VALUE : 0x2,
		KEY_CREATE_SUB_KEY : 0x4,
		KEY_ENUMERATE_SUB_KEYS : 0x8,
		KEY_NOTIFY : 0x10,
		KEY_CREATE_LINK : 0x20,
		DELETE : 0x10000,
		READ_CONTROL : 0x20000,
		WRITE_DAC : 0x40000,
		WRITE_OWNER : 0x80000,

		con : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control",
		ses : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Session Manager",
		env : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Session Manager\\Environment",
		uenv : "HKEY_CURRENT_USER\\System\\CurrentControlSet\\Control\\Session Manager\\Environment",
		mem : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Session Manager\\Memory Management",
		win : "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon",
		cur : "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion",
		svc : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services",
		// Act as a NT server "NeutralizeNT4Emulator"=dword:00000001 (must restart netlogon)
		svc_par : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\Netlogon\\Parameters",
		run : "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
		once : "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce",
		prn : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Print\\Printers",
		prn_port : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports",
		prn_drv3 : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Control\\Print\\Environments\\Windows NT x86\\Drivers\\Version-3",
		prn_ins : "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Print\\Providers\\LanMan Print Services\\Servers",
		pro : "HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\Installer\\Products",
		// Windows 2000/XP - http://www.winguides.com/registry/display.php/1298
		// Policy=(Ignore=00, Warn=01Block=02) Windows 2003
		drv_sign1 : "HKEY_LOCAL_MACHINE\\Software\\Policies\\Microsoft\\Windows NT\\Driver Signing",
		drv_sign2 : "HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Driver Signing",
		lan_par : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\lanmanserver\\parameters",
		// \\{605DA0E2-2EB2-4ECA-89F0-4EDDAD382641}\\NameServers
		tcpip : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\Tcpip\\Parameters",
		// \\{605DA0E2-2EB2-4ECA-89F0-4EDDAD382641}\\NameServers
		tcpip_int : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\Tcpip\\Parameters\\Interfaces",
		lan_par : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\lanmanserver\\parameters", // srvcomment
		nbt_int : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\NetBT\\Parameters\\Interfaces", // srvcomment
		//IE
		ie_cu : hkcu + "\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings",
		ie_lm : "HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings",
		// http://support.microsoft.com/default.aspx?scid=kb__H;en-us;Q175500
		ie_styles : hkcu + "\\Software\\Microsoft\\Internet Explorer\\Styles", // MaxScriptStatements
		// SNA
		sna_par : "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services\\SNABase\\Parameters",
		sna_netsoft : "HKEY_LOCAL_MACHINE\\SOFTWARE\\NetSoft\\NS/Elite\\CurrentVersion\\Workspaces\\Mainframe Workspace\\Links\\TN3270\\NS/Router 3270", // Data (BINARY)
		
		usr_shell : "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders",	
		ore_url : /^http:\/\/.+\..+|^ftp:\/\/.+\..+/ig
	}
	
	/////////////////////////////////////
	//// DEFAULT
		
	var o_options = {
		
	}
	
	var initialize = function initialize(bForce){
		if(b_initialized && !bForce) return;

		d_ipconfig = new ActiveXObject("Scripting.Dictionary")
		o_this.resource = oWno.ComputerName
		o_this.computer = oWno.ComputerName
		o_this.userdomain = oWno.UserDomain
		o_this.username = oWno.UserName
		o_this.domainuser = o_this.userdomain + "\\" + o_this.username		
		o_this.password = null
		o_this.isAdmin()
		
		b_initialized = true
	}
	
	this.setOptions = function setOptions(oOptions){
		if(__H.isObject(oOptions)) {
			return Object.extend(o_options,oOptions,true)
		}
		return false
	}
	
	this.getOptions = function(){
		return o_options
	}
	
	/////////////////////////////////////
	//// 
	
			
	this.isAdmin = function isAdmin(){
		this.isadmin = (oWsh.Run("%comspec% /c dir /b \\\\%computername%\\admin$\\notepad.*",0,true) == 0)
		return this.isadmin
	}

	this.setSystem = function setSystem(sComputer,sUser,sPass){
		initialize()
		try{
			this.computer = !__H.isStringEmpty(sComputer) ? sComputer : this.computer
			this.userdomain = this.computer
			this.password = !__H.isStringEmpty(sPass) ? sPass : this.password

			if(!__H.isStringEmpty(sUser)){
				if(sUser.match(/(.+)\\(.+)/g)) this.userdomain = RegExp.$1, this.username = RegExp.$2
				else this.username = sUser
				this.domainuser = this.userdomain + "\\" + this.username
			}
			else if(!this.isLocalHost()){
				__HExp.SYSRemoteAccessDenied("User name not defined for remote system: " + sComputer)
			}
			else {
				this.username = null
				this.password = null
				this.userdomain  = null
			}
			this.resource = this.computer
		}
		catch(ee){alert(gg)
			__HLog.error(ee,this)
			return false;
		}
	}

	this.isLocalHost = function isLocalHost(){
		return (this.computer.toUpperCase() == oWno.ComputerName.toUpperCase());
	}

	this.isLocalHost2 = function isLocalHost2(sComputer){
		try{
			if(typeof(sComputer) != "string") return false;
			else if(sComputer.toUpperCase() == oWno.ComputerName.toUpperCase() || sComputer == "127.0.0.1") return true;
			else if(this.isIPAddress(sComputer)){
				return (oWsh.Run("%comspec% /c ping -a " + sComputer + " -n 1 -w 50 | find /i \"%computername%\"",__HIO.hide,true) == 0);
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
	
	this.run = function run(sCommand,bShow){
		if(bShow) oWsh.Run("%comspec% /k " + sCommand,__HIO.show,false)
		else oWsh.Run("%comspec% /c " + sCommand,__HIO.hide,false)
	}

	this.timeStop = function timeStop(ID){
		try{
			if(ID != null){
				clearTimeout(ID);
				clearInterval(ID)
				ID = null
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.sleep = function sleep(i){
		__HUtil.sleep(i);
	}

	this.regActiveX = function regActiveX(sName,sFile,bRegistered){
		try{
			try{
				return new ActiveXObject(sName);
			}
			catch(eee){
				if(!bRegistered && oFso.FileExists(sFile)){
					if(this.regsvr32(sFile,false)){
						return this.regActiveX(sName,sFile,true);
					}
				}
				else throw new Error(__HLog.errorCode("error"),"Unable to register DLL file: " + sFile);
				return false;
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.regsvr32 = function regsvr32(sFile,bUnRegister){
		try{
			if(oFso.FileExists(sFile = oFso.GetAbsolutePathName(sFile))){
				var sMode = !bUnRegister ? " " : " /u ";
				return (oWsh.Run("%comspec% /c regsvr32 /s" + sMode + "\"" + sFile + "\"",__HIO.hide,true) == 0);
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.ping = function ping(iCount){
		try{
			if(this.isLocalHost()) return true
			iCount = !isNaN(iCount) ? iCount : 1
			if(oWsh.Run("%comspec% /c ping " + this.computer + " -n " +  iCount + " -l 100 -w 1 -f -i 1 | find /i \"TTL\" >nul",__HIO.hide,true) != 0){
				__HLog.debug("Unable to ping() system: " + this.computer)
				return false
			}
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.isMacAddress = function isMacAddress(s){
		try{
			if(__H.isStringEmpty(s)) return false;
			else if(s.length == 16){ //
				return s.isSearch(/[a-z0-9]{2}-[a-z0-9]{2}-[a-z0-9]{2}-[a-z0-9]{2}-[a-z0-9]{2}-[a-z0-9]{2}$/ig)
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.isIPAddress = function isIPAddress(s){
		try{
			var bResult = false, a;
			var oRe = new RegExp("([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})","g")

			if(__H.isStringEmpty(s)) return false;
			else if(a = s.isSearch(oRe)){
				if(s.length == a[0].length){
					bResult = true;
					for(var i = 1; i < 5; i++){
						if(parseInt(a[i]) > 255) bResult = false;
					}
				}
			}
			return bResult;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.systeminfo = function systeminfo(){
		this.netUseIPC()
		this.run("systeminfo /s " + this.computer + " | find /v /i \"File 1\">%temp%\\systeminfo.log && start \"Must Title\" %temp%\\systeminfo.log")
	}

	this.sysinfo = function sysinfo(){
		try{
			return {
				Computer	: this.computer,
				OS			: oWsh.RegRead(__HKeys.env + "\\OS"),
				OSversion	: oWsh.RegRead(__HKeys.cur + "\\CurrentVersion"),
				ServicePack	: oWsh.RegRead(__HKeys.cur + "\\CSDVersion"),
				Username	: oWsh.RegRead(__HKeys.win + "\\DefaultUserName"),
				Domain		: oWsh.RegRead(__HKeys.win + "\\DefaultDomainName"),
				InstallDate	: oWsh.RegRead(__HKeys.cur + "\\InstallDate")
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}	

	this.ipaddress = function ipaddress(){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			if(this.isIPAddress(this.computer)) return this.computer
			else {
				var sCmd = "%comspec% /c nbtstat -R >nul & ping " + this.computer + " -n 1 -l 32 | find /i \"TTL\" 2>nul 1>" + sFile;
				if(oWsh.Run(sCmd,__HIO.hide,true) == 0 && oFso.FileExists(sFile)){
					var oFile = oFso.OpenTextFile(sFile,__HIO.read,true,__HIO.TristateUseDefault);
					var s = (oFile.ReadAll()).trim() // Reply from 10.35.5.12: bytes=32 time=30ms TTL=125
					oFile.Close()
					if(this.isIPAddress(s.replace(/.+ ([0-9.]{7,15}): .+/,"$1"))){
						return RegExp.$1 + "." + RegExp.$2 + "." + RegExp.$3 + "." + RegExp.$4;
					}
				}
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.ipnslookup = function ipnslookup(sDNSServer){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			if(this.isIPAddress(this.computer)) return this.computer
			else {
				sDNSServer = !__H.isStringEmpty(sDNSServer) ? sDNSServer : ""
				var sCmd = "%comspec% /c ipconfig /flushdns >nul & nslookup -type=A " + this.computer + " " + sDNSServer + " | find /v /i \"" + sDNSServer + "\" | find /i \"address\">" + sFile;
				if(oWsh.Run(sCmd,__HIO.hide,true) == 0 && oFso.FileExists(sFile)){
					var oFile = oFso.OpenTextFile(sFile,__HIO.read,true,__HIO.TristateUseDefault);
					var s = (oFile.ReadAll()).trim()
					oFile.Close()
					if(this.isIPAddress(s)){
						return RegExp.$1 + "." + RegExp.$2 + "." + RegExp.$3 + "." + RegExp.$4;
					}
				}
			}
			return this.ipaddress();
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.ipconfig = function ipconfig(){
		try{
			if(d_ipconfig.Exists(this.computer)) return d_ipconfig(this.computer)
			var o = {}, aKeys
			
			if(__H.isUndefined(__HReg)){
				throw new Error(this.errorCode("error"),"Registry Class Library needs to be loaded!")
			}
			if(!__HReg.initialize(this.computer,this.domainuser,this.password,true)){
				return o
			}
						
			o.IpAddress = "0.0.0.0"
			o.DNSSearchList = __HReg.getValue(__HKeys.tcpip,"SearchList")
			o.Domain = __HReg.getValue(__HKeys.tcpip,"Domain")
			o.DhcpNameServer = __HReg.getValue(__HKeys.tcpip,"DhcpNameServer")
			o.Hostname = (__HReg.getValue(__HKeys.tcpip,"Hostname")).toUpperCase()
			o.IPRouter = !!__HReg.getValueDWORD(__HKeys.tcpip,"IPEnableRouter")
			o.ReservedPorts = __HReg.getValueMULTI(__HKeys.tcpip,"ReservedPorts")
			
			if(aKeys = __HReg.getKeys(__HKeys.tcpip + "\\DNSRegisteredAdapters")){
				for(var j = 0, iLen = aKeys.length; j < iLen; j++){
					if(aKeys[j].substring(0,1) == "{"){ // {0711509E-13F3-412A-BD03-8962B397D622}
						o.EnableDHCP = !!__HReg.getValueDWORD(__HKeys.tcpip_int + "\\" + aKeys[j],"EnableDHCP")
						o.IpAddress = __HReg.getValueMULTI(__HKeys.tcpip_int + "\\" + aKeys[j],(o.EnableDHCP ? "DhcpIPAddress" : "IPAddress"))
						o.PrimaryDomainName = __HReg.getKeys(__HKeys.tcpip + "\\DNSRegisteredAdapters\\" + aKeys[j])
						o.NameServers = __HReg.getValue(__HKeys.tcpip_int + "\\" + aKeys[j],"NameServer") (o.EnableDHCP ? "Dhcp" : "")
						o.SubnetMask = __HReg.getValue(__HKeys.tcpip_int + "\\" + aKeys[j],(o.EnableDHCP ? "DhcpSubnetMask" : "SubnetMask"))
						o.DhcpServer = __HReg.getValue(__HKeys.tcpip_int + "\\" + aKeys[j],"DhcpServer")
						o.DefaultGateway = __HReg.getValue(__HKeys.tcpip_int + "\\" + aKeys[j],(o.EnableDHCP ? "DhcpDefaultGateway" : "DefaultGateway"))							
						o.LeaseObtained = __HReg.getValueDWORD(__HKeys.tcpip_int + "\\" + aKeys[j],"LeaseObtainedTime")
						o.LeaseTerminate = __HReg.getValueDWORD(__HKeys.tcpip_int + "\\" + aKeys[j],"LeaseTerminatesTime")
						o.WINSservers = __HReg.getValueMULTI(__HKeys.nbt_int + "\\Tcpip_" + aKeys[j],"NameServerList")
					}
				}
			}

			d_ipconfig.Add(this.computer,o)
			return o;
		}
		catch(ee){
			__HLog.error(ee,this)
			return null;
		}
	}

	this.ipnetmask = function ipnetmask(s){
		try{
			if(this.isIPAddress(s) || (s = this.ipaddress(this.computer))){
				var a = s.split(".");
				return a[0] + "." + a[1] + "." + a[2] + ".0";
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.ipcomputer = function ipcomputer(sIpAddress){
		try{
			var sComputer = false;
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			if(!this.isIPAddress(sIpAddress) && !(sIpAddress = this.ipaddress(sIpAddress))) return false
			//else if(!this.ping(sIpAddress)) return false
			var sCmd = "%comspec% /c nbtstat -R >nul & nbtstat -A " + sIpAddress + " | find /i \"<03>\" 2>nul 1>" + sFile;
			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,true,__HIO.TristateUseDefault);
				var oRe = new RegExp("[ \t]+([a-z0-9\-]+)[ \t]+.*","i"); //	MAC Address = 00-B0-D0-D5-EA-64
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oRe.exec(sLine)){
						sComputer = (RegExp.$1).toUpperCase();
						break;
					}
				}
				oFile.Close();
			}
			return sComputer;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.ipcomputer2 = function ipcomputer2(sIpAddress){
		try{
			var sComputer = false;
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			var sCmd = "%comspec% /c nbtstat -R >nul & ping -a -n 1 -l 32 " + sIpAddress + " | find /i \"pinging\" 2>nul 1>" + sFile;
			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				var oRe = new RegExp("pinging[ \t]+([a-z0-9\-\.]+)[ \t].+","ig"); // Pinging kawari.dnb.no [10.20.16.28] with 32 bytes of data:
				sComputer = (oFile.ReadAll()).trim()
				oFile.Close();
				sComputer = sComputer.match(oRe) ? (RegExp.$1).toLowerCase() : false
			}
			return sComputer;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.macaddress = function macaddress(bSeparator){
		try{
			if(this.ping()) return false
			var s = false
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			var sCmd = "%comspec% /c nbtstat -a " + this.computer + " | find /i \" = \" > " + sFile;
			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				var oRe = new RegExp(".*(= )([a-z0-9\-]{17}).*","i"); //	MAC Address = 00-B0-D0-D5-EA-64
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oRe.exec(sLine)){
						s = (RegExp.$2).toUpperCase();
						s = !bSeparator ? s.replace(/[\-]*/g,"") : s;
						break;
					}
				}
				oFile.Close();
			}
			return s;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netSession = function netSession(sPsExec){
		try{
			var a = [];
			if(this.ping()) return a
			
			var sFile = __HIO.temp + "\\" + oFso.GetTempName(), sCmd
			sPsExec = oFso.FileExists(sPsExec) ? "\"" + sPsExec + "\"" : "psexec.exe"
			if(!this.isLocalHost()) sCmd = "%comspec% /c " + sPsExec + " \\\\" + this.computer + " net session | find /i \"Windows\" > " + sFile;
			else sCmd = "%comspec% /c net session | find /i \"Windows\" > " + sFile;
			
			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				var oRe = new RegExp("\\\\([0-9\.]+)[ \s]+([a-z0-9]+)[ \s]+([a-z0-9 ]+) ([0-9])[ \s]+([0-9:]+).*","i"); //	MAC Address = 00-B0-D0-D5-EA-64
				while(!oFile.AtEndOfStream){
					var sLine = oFile.ReadLine();
					if(oRe.exec(sLine)){
						var o = {
							computer   : RegExp.$1,
							username   : RegExp.$2,
							clienttype : RegExp.$3,
							opens      : RegExp.$4,
							idletime   : RegExp.$5
						}
						a.push(o);
					}
				}
				oFile.Close();
			}
		}
		catch(ee){
			__HLog.error(ee,this)
		}
		return a;
	}

	this.netshDomain = function netshDomain(sDHCPServer){
		try{
			var sDomain = false, sNetMask
			if(sNetMask = this.ipnetmask()){
				var sFile = __HIO.temp + "\\" + oFso.GetTempName()
				var sCmd = "%comspec% /c netsh dhcp server \\\\" + sDHCPServer + " scope " + sNetMask + " dump | find /i \"dhcp\" | find /i \"optionvalue 15\" >" + sFile
				oWsh.Run(sCmd,__HIO.hide,true)
				// Dhcp Server 10.59.252.5 Scope 10.59.204.0 set optionvalue 15 STRING "Mosjoen.helgelandsb.lan"
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				var s = (oFile.ReadAll()).trim()
				oFile.Close();
				var oRe = new RegExp(".+ STRING \"([a-z\.]{6,})\"","ig");
				if(oRe.exec(s)){
					sDomain = RegExp.$1
					if(sDomain.match(/([a-z]+)\.([a-z]+)\.([a-z]{2,5})/ig)) sDomain = RegExp.$1; // Mosjoen
				}
			}
			return sDomain.capitalize();
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netshComputer = function netshComputer(sName){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			var sWINS = (this.ipconfig()).WINSservers[0]

			oWsh.Run("%comspec% /c nbtstat -R>nul & nbtstat -RR>nul & ipconfig /flusgdns>nul",__HIO.hide,true)
			var sCmd = "%comspec% /c for /f \"tokens=4\" %a in ('netsh wins server \\\\" + sWINS + " show name " + sName
						+ " 03 ^| find /i \"IP Address\"') do for /f \"tokens=2\" %i in ('nslookup %a ^| find /i \"Name:\"') do echo %i>" + sFile

			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				var s = (oFile.ReadAll()).trim() // sm78273.dnbnor.net
				oFile.close()
				return (s.replace(/([a-z0-9\-_]+).*$/ig,"$1")).toUpperCase()
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netshIPAddress = function netshIPAddress(sName){
		try{
			var sFile = __HIO.temp + "\\" + oFso.GetTempName()
			var sWINS = (this.ipconfig()).WINSservers[0]
			oWsh.Run("%comspec% /c nbtstat -R>nul & nbtstat -RR>nul & ipconfig /flusgdns>nul",__HIO.hide,true)

			/*
			***You have Read and Write access to the server 10.68.68.68***

			Name                  : AB37360        [03h]
			NodeType              : 3
			State                 : ACTIVE
			Expiration Date       : 18. oktober 2006 08:03:06
			Type of Rec           : UNIQUE
			Version No            : 0 3b7ad5f
			RecordType            : DYNAMIC
			IP Address            : 10.195.4.83
			Command completed successfully.
			*/
			var sCmd = "%comspec% /c for /f \"tokens=4\" %a in ('netsh wins server \\\\" + sWINS + " show name " + sName
						+ " 03 ^| find /i \"IP Address\"') do echo %a>" + sFile
			if(oWsh.Run(sCmd,__HIO.hide,true) == 0){
				var oFile = oFso.OpenTextFile(sFile,__HIO.read,false,__HIO.TristateUseDefault);
				oFile.close()
				return (oFile.ReadAll()).trim() // 10.195.4.83
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.launch = function launch(sOpt){
		try{
			if(sOpt == "msinfo32"){
				oWsh.Run("winmsd /computer " + this.computer,__HIO.hide,false);
			}
			else if(sOpt == "eventvwr"){
				oWsh.Run("%systemroot%\\system32\\eventvwr.exe \\\\" + this.computer,__HIO.hide,false)
			}
			else if(sOpt == "compmgmt"){
				oWsh.Run("%systemroot%\\system32\\compmgmt.msc -s /computer:\\\\" + this.computer,__HIO.hide,false)
			}
			else if(sOpt == "explorer"){
				oWsh.Run("%comspec% /c start \\\\" + this.computer + "\\c$",__HIO.hide,false)
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.bytes = function bytes(iBytes){
		try{
			if(typeof(iBytes = parseInt(iBytes)) != "number") return 0
			if(iBytes >= 1000000000) return parseFloat(iBytes/1047000000).toDecimal(2) + " GB";
			else if(1000000000 > iBytes && iBytes >= 1000000) return parseFloat(iBytes/1047000).toDecimal(2) + " MB";
			else if(1000000 > iBytes && iBytes >= 1000) return parseFloat(iBytes/1024).toDecimal(2) + " KB";
			else return parseFloat(iBytes) + " Bytes";
		}
		catch(ee){
			__HLog.error(ee,this)
			return 0;
		}
	}

	this.netLocalgroup = function netLocalgroup(sPsExec,sLocalGroup,sGlobalGroup){
		try{
			if(__H.isStringEmpty(sLocalGroup,sGlobalGroup,sPsExec)) return false
			sPsExec = oFso.FileExists(sPsExec) ? "\"" + sPsExec + "\"" : "%comspec% /c psexec.exe"
			var r, sCmd = sPsExec + " \\\\" + this.computer + " -u " + this.domainuser + " -p " + this.password + " net localgroup " + sLocalGroup + " " + sGlobalGroup + " /ADD"
			if((r=oWsh.Run(sCmd,__HIO.hide,true)) != 0 && r != 2) return false; // r=0 update OK. r=2: Already updated
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netUseIPC = function netUseIPC(){
		try{
			this.netUseDel();
			var sCmd1 = "%comspec% /c net use \\\\" + this.computer + "\\IPC$ /u:" + this.domainuser + " " + this.password;
			var sCmd2 = "%comspec% /c net use \\\\" + this.computer + "\\IPC$";
			return ((oWsh.Run(sCmd1,__HIO.hide,true) != 0) || (oWsh.Run(sCmd2,__HIO.hide,true) != 0))
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netUseAdmin = function netUseAdmin(){
		try{
			if(this.isLocalHost()) return true;
			if(!this.ping()) return false
			var sCmd1 = "%comspec% /c net use \\\\" + this.computer + "\\admin$";
			var sCmd2 = "%comspec% /c net use \\\\" + this.computer + "\\admin$ /u:" + this.domainuser + " " + this.password;
			return ((oWsh.Run(sCmd1,__HIO.show,true) == 0) || (oWsh.Run(sCmd2,__HIO.show,true) == 0));
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netUseDel = function netUseDel(bAll){
		try{
			if(bAll) oWsh.Run("%comspec% /c net use * /del /y",__HIO.hide,false);
			else {
				oWsh.Run("%comspec% /c net use \\\\" + this.computer + "\\admin$ /del /y",__HIO.hide,true);
				oWsh.Run("%comspec% /c net use \\\\" + this.computer + "\\IPC$ /del /y",__HIO.hide,false);
			}
			return true;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.netadmin = function netadmin(){
		try{
			if(this.isLocalHost()) return true;
			if(this.ping()){
				if(oWsh.Run("%comspec% /c dir /ad /b \\\\" + this.computer + "\\c$ >nul",__HIO.hide,true) != 0){
					var up = !__H.isStringEmpty(this.username,this.password) ? " /u:" + this.domainuser + " " + this.password : ""
					var sCmd = "%comspec% /c mode con cols=80 lines=10 & color 87 & Title Admin connection on: " + this.computer + " & net use \\\\" + this.computer + "\\admin$ " + up + " /p:no"
					if(oWsh.Run(sCmd,__HIO.show,true) != 0){
						return false
					}
				}
				return true
			}
			return false;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.systemfolder = function systemfolder(iDir,iPversion){
		try{
			var sDir = false;
			/*
			http://msdn.microsoft.com/library/default.asp?url=/library/en-us/graphics/hh/graphics/prtinst_976v.asp
			66000 = C:\WINNT\System32\spool\DRIVERS\W32X86\2
			66001 = C:\WINNT\System32\spool\PRTPROCS\W32X86\2
			66002 = system32 (always)
			66003 = C:\WINNT\System32\spool\PRTPROCS\W32X86\2

			http://msdn.microsoft.com/library/default.asp?url=/library/en-us/infguide/hh/infguide/infguide_5ew5.asp
			00 Null LDID. This LDID can be used to create a new LDID
			01 Source Drive:\pathname
			10 Machine directory (Maps to the Windows directory on a Server-Based Setup.)
			11 System directory
			12 IOSubsys directory
			13 Command directory
			17 INF Directory
			18 Help directory
			20 Fonts
			21 Viewers
			22 VMM32
			23 Color directory
			24 Root of drive containing the Windows directory
			25 Windows directory
			26 Guarenteed boot device for Windows (Winboot)
			28 Host Winboot
			30 Root directory of the boot drive, Ex: DefaultDestDirs=30,bin ; Direct copies go to Boot:\bin
			31 Root directory for Host drive of a virtual boot drive
			*/
			switch(iDir){
				case '66000' : case '52' : {
					sDir = (iPversion == 2) ? __HIO.w32 : __HIO.w32_2
					break;
				}
				case '66001' : {
					sDir = __HIO.w32;
					break;
				}
				case '11' : case '12' : case '66002' : {
					sDir = __HIO.system32;
					break;
				}
				case '17' : { sDir = __HIO.windir + "\\inf"; break;}
				case '18' : { sDir = __HIO.windir + "\\help"; break;}
				case '25' : { sDir = __HIO.windir; break;}
				default : break;
			}
			return sDir;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.msdos = function msdos(sCommand,iMaxwait){
		try { // Ex sCommand: %comspec% /c dir
			iMaxwait = __H.isNumber(iMaxwait) ? iMaxwait : 30000; // 30 seconds
			var iSleep = 3
			var o = {};
			o.ExitCode = 9999
			var oExec = oWsh.Exec(sCommand);
			o.ProcessID = oExec.ProcessID; // (Read-Only) the process ID (PID) of the run process
			for(var iWait = iSleep; true; iWait += iSleep){
				if(oExec.Status == 0 || iWait > iMaxwait){
					break;
				}
				__HUtil.sleep(iSleep);
			}
			o.ExitCode = oExec.ExitCode; // (Read-Only) exit code returned by the run process
			o.message = (oExec.StdOut.ReadAll()).trim()
			oExec.Terminate()
			o.message_err = (oExec.StdErr.ReadAll()).trim()

			return o;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.reboot = function reboot(){
		try{
			if(__H.isUndefined(__HWMI)){
				throw new Error(this.errorCode("error"),"WMI Library needs to be loaded!")
			}
			if(__HWMI.setServiceWMI(this.computer,"root/cimv2",this.domainuser,this.password)){
				var oService = __HWMI.setServiceWMI()			
			}
			else return false
			var oColItems = oService.ExecQuery("Select * from Win32_OperatingSystem where Primary=true","WQL",48);
			for(var oEnum = new Enumerator(oColItems); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				oItem.Reboot();
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

var __HSys = new __H.Sys()
var __HKeys	= __HSys.reg_keys