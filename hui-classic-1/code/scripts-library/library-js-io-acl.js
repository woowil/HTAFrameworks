// nOsliw Solutions - HTML Aplication Framework.
//**Start Encode**

/*

File:     library-js-io-acl.js
Purpose:  Development script
Author:   Woody Wilson
Created:  08.07.2002
Version:  
Dependency:  library-js.js, library-js-io-file.js, library-js-io-folder.js, library-js-system.js

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
	if(!oFso || !oWsh || !oWno || !oXml){
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

var oAcl = new io_acl_object();

////////////////////////////////////////////////////////////////////////////////////////////////
/////// ACL FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////


function io_acl_object(){ //http://www.serverwatch.com/tutorials/article.php/1476871
	
	this.ADS_ACETYPE_ACCESS_ALLOWED	= 0
	this.ADS_ACETYPE_ACCESS_DENIED	= 1
	this.ADS_ACEFLAG_INHERIT_ACE	= 2
	this.ADS_ACEFLAG_SUB_NEW	= 9
	
	// Share right
	this.SHARE_RIGHT_READ			= 1179817
	this.SHARE_RIGHT_FULL			= 2032127
	this.SHARE_RIGHT_CHANGE			= 1245631
	this.SHARE_RIGHT_DENY			= 131241
	
	// ACE to specified DACL
	this.ACE_RIGHT_READ			= 0x80000000
	this.ACE_RIGHT_EXECUTE			= 0x20000000
	this.ACE_RIGHT_WRITE			= 0x40000000
	this.ACE_RIGHT_DELETE			= 0x10000
	this.ACE_RIGHT_FULL			= 0x10000000
	this.ACE_RIGHT_CHANGE_PERMS		= 0x40000
	this.ACE_RIGHT_TAKE_OWNERSHIP		= 0x80000
	
	this.ADS_RIGHT_DELETE                 = 0x10000
	this.ADS_RIGHT_READ_CONTROL           = 0x20000
	this.ADS_RIGHT_WRITE_DAC              = 0x40000
	this.ADS_RIGHT_WRITE_OWNER            = 0x80000
	this.ADS_RIGHT_SYNCHRONIZE            = 0x100000
	this.ADS_RIGHT_ACCESS_SYSTEM_SECURITY = 0x1000000
	this.ADS_RIGHT_GENERIC_READ           = 0x80000000
	this.ADS_RIGHT_GENERIC_WRITE          = 0x40000000
	this.ADS_RIGHT_GENERIC_EXECUTE        = 0x20000000
	this.ADS_RIGHT_GENERIC_ALL            = 0x10000000
	this.ADS_RIGHT_DS_CREATE_CHILD        = 0x1
	this.ADS_RIGHT_DS_DELETE_CHILD        = 0x2
	this.ADS_RIGHT_ACTRL_DS_LIST          = 0x4
	this.ADS_RIGHT_DS_SELF                = 0x8
	this.ADS_RIGHT_DS_READ_PROP           = 0x10
	this.ADS_RIGHT_DS_WRITE_PROP          = 0x20
	this.ADS_RIGHT_DS_DELETE_TREE         = 0x40
	this.ADS_RIGHT_DS_LIST_OBJECT         = 0x80
	this.ADS_RIGHT_DS_CONTROL_ACCESS      = 0x100
	
	// Access Control Entry Type Values
 	// Possible values for the IADsAccessContronEntry::AceType property
	this.ADS_ACETYPE_ACCESS_ALLOWED           = 0
	this.ADS_ACETYPE_ACCESS_DENIED            = 0x1
	this.ADS_ACETYPE_SYSTEM_AUDIT             = 0x2
	this.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT    = 0x5
	this.ADS_ACETYPE_ACCESS_DENIED_OBJECT     = 0x6
	this.ADS_ACETYPE_SYSTEM_AUDIT_OBJECT      = 0x7
	
	// Access Control Entry Inheritance Flags
	// Possible values for the IADsAccessControlEntry::AceFlags property
	this.ADS_ACEFLAG_UNKNOWN                  = 0x1 // 
	this.ADS_ACEFLAG_INHERIT_ACE              = 0x2 // child objects will inherit ACE of current object
	this.ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE = 0x4 // prevents ACE inherited by the object from further propagation 
	this.ADS_ACEFLAG_INHERIT_ONLY_ACE         = 0x8 // indicates ACE used only for inheritance (it does not affect permissions on object itself)
	this.ADS_ACEFLAG_INHERITED_ACE            = 0x10 // indicates that ACE was inherited
	this.ADS_ACEFLAG_INHERIT_FLAGS      	  = 0x1F // indicates that inherit flags are valid (provides confirmation of valid settings)
	this.ADS_ACEFLAG_SUCCESSFUL_ACCESS        = 0x40 // for auditing success in system audit ACE
	this.ADS_ACEFLAG_FAILED_ACCESS            = 0x80 // for auditing failure in system audit ACE
	
	// Registry Permission Type Values
	this.KEY_QUERY_VALUE 		= 0x0001
	this.KEY_SET_VALUE 		= 0x0002
	this.KEY_CREATE_SUB_KEY 	= 0x0004
	this.KEY_ENUMERATE_SUB_KEYS 	= 0x0008
	this.KEY_NOTIFY 		= 0x0010
	this.KEY_CREATE_LINK 		= 0x0020
	this.DELETE 			= 0x00010000
	this.READ_CONTROL 		= 0x00020000
	this.WRITE_DAC 			= 0x00040000
	this.WRITE_OWNER 		= 0x00080000

}

function io_acl_security(sOpt,sFileFolder,sPermissions,bNotSetDescriptor,sOpt5,sOpt6){
	var oADs, oAclSD, oDAcl
	try{ // see http://www.visualbasicscript.com/topic.asp?TOPIC_ID=2220 http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B279682
		if(!js_str_isdefined(sOpt,sFileFolder)) return false
		else if(!js.adssecurityx) return false // ADsSecurity.dll (ADsSecurity) must be loaded and associated globally as js.adssecurityx
		else if(!oFso.FolderExists(sFileFolder) && !oFso.FileExists(sFileFolder)) return false
		//else if(!io_folder_readable(sFileFolder) && !io_file_readable(sFileFolder)) return false
		
		oADs = js.adssecurityx, sResult = true;
		oAclSD = oADs.GetSecurityDescriptor("FILE://" + sFileFolder)
		oDAcl = oAclSD.DiscretionaryAcl;
		var bSetDescriptor = bNotSetDescriptor ? false : true
		
		if(sOpt.match(/getdacl/ig)){
			return oDAcl;
		}
		else if(sOpt.match(/replace/ig)){ // Removes all existing aces from dacl first
			for(var oItem, oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				oDAcl.RemoveAce(oItem);
			}
		}
		else if(sOpt.match(/ntauthority/ig)){ // For some reason if ace includes "NT AUTHORITY" then existing ace does not get readed to dacl
			var oRe = /nt authority\\(.*)/ig
			for(var oItem, oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if((oItem.Trustee).match(oRe)){
					oItem.Trustee = RegExp.$1
				}
			}
		}
		else if(sOpt.match(/getowner/ig)){
			return oAclSD.Owner;
		}
		else if(sOpt.match(/getacls$/ig)){ // htm_div_html("settable",io_acl_security("showacls","C:\\temp"))
			var aAcl = new Array()
			for(var oItem, oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.AceType == 0){
					var o = new Object()
					o.Trustee = oItem.Trustee // BUILTIN\Administrators
					o.AccessMask = oItem.AccessMask // 1245631
					o.Rights = io_acl_getrights(oItem.AccessMask); // CHANGE or FULL CONTROL or READ or WRITE
					o.AceFlags = oItem.AceFlags
					o.Flags = io_acl_getflags(oItem.AceFlags);
					o.AceType = oItem.AceType
					aAcl.push(o);
				}
			}
			return aAcl;
		}
		else if(sOpt.match(/getacls_simple/ig)){
			var aAcl = new Array()
			for(var oItem, oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if(oItem.AceType == 0 && oItem.AccessMask >= 0){ // Ignore duplicates
					var o = new Object()
					o.Trustee = oItem.Trustee // BUILTIN\Administrators
					o.Rights = io_acl_getrights(oItem.AccessMask); // CHANGE or FULL CONTROL or READ or WRITE
					if(typeof(o.Rights) == "string") aAcl.push(o); // ignore undefined
				}
			}
			return aAcl;
		}
		else if(sOpt.match(/showacls$/ig)){
			var s = "Trustee; AccessMask; Rights; AceFlags; Flags; AceType", a
			if(a = io_acl_security("getacls",sFileFolder)){
				for(var i = 0, len = a.length; i < len; i++){
					s = s + "\n" + a[i].Trustee + ";" + a[i].AccessMask + ";" + a[i].Rights + ";" + a[i].AceFlags + ";" + a[i].Flags + ";" + a[i].AceType;
				}
			}
			return s
		}
		else if(sOpt.match(/showacls_simple/ig)){
			var s = "Trustee; Rights", a
			if(a = io_acl_security("getacls_simple",sFileFolder)){
				for(var i = 0, len = a.length; i < len; i++){
					s = s + "\n" + a[i].Trustee + ";" + a[i].Rights;
				}
			}
			return s
		}
		else if(sOpt.match(/setfile|setfolder/ig)){
			var sSplit1 = sOpt5 ? sOpt5 : "::"
			var sSplit2 = sOpt6 ? sOpt6 : ":"
			if(!js_str_isdefined(sPermissions)) return false
			var aPerms = io_acl_getpermissions(sPermissions,sSplit1,sSplit2);
			var oRe = /[a-z ]*\\{0,1}([a-z ]+)/ig
			for(var i = 0, len = aPerms.length; i < len; i++){
				var oPerm = aPerms[i];
				// delete ACE's belonging to user about to add an ace for
				for(var oItem, oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
					oItem = oEnum.item();
					var sName = ((oItem.Trustee).match(oRe) ? RegExp.$1 : oItem.Trustee).toLowerCase();
					if(sName == oPerm.name){
						oDAcl.RemoveAce(oItem)
					}
				}
				// if action is to del ace then following clause skips addace
				if(oPerm.type == "add"){
					var b1, b2
					if(sOpt.match(/setfolder/ig)){
						b1 = io_acl_securitydacl("setacentfs",oDAcl,oPerm.name,oPerm.perm,oAcl.ADS_ACETYPE_ACCESS_ALLOWED,oAcl.ADS_ACEFLAG_SUB_NEW);
						b2 = io_acl_securitydacl("setacentfs",oDAcl,oPerm.name,oPerm.perm,oAcl.ADS_ACETYPE_ACCESS_ALLOWED,oAcl.ADS_ACEFLAG_INHERIT_ACE);
					}
					else b1 = io_acl_securitydacl("setacentfs",oDAcl,oPerm.name,oPerm.perm,oAcl.ADS_ACETYPE_ACCESS_ALLOWED,0), b2 = true;
					if(!b1 || !b2) sResult = false
				}
			}
			io_acl_security("ntauthority",sFileFolder);
			if(bSetDescriptor) io_acl_securitydacl("setdacl",oDAcl,sFileFolder);
		}
		else return false		
		
		return sResult;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		js_str_kill(oADs,oAclSD);
	}
}

function io_acl_securitydacl(sOpt,oDAcl,sFileFolder,sTrustee,sPerm,iAceType,iAceFlag){
	var oADs,oAce,oDAcl
	try{
		if(!js.adssecurityx) return false // ADsSecurity.dll (ADsSecurity) must be loaded and associated globally as js.adssecurityx
		oADs = js.adssecurityx		
		oAclSD = oADs.GetSecurityDescriptor("FILE://" + sFileFolder)
		
		if(sOpt.match(/setaceshare/ig)){
			oDAcl = oDAcl ? oDAcl : oAclSD.DiscretionaryAcl;
			oAce = new ActiveXObject("AccessControlEntry");
			oAce.Trustee = sTrustee ? sTrustee : "EVERYONE"
			oAce.AccessMask = sPerm ? sPerm : oAcl.SHARE_RIGHT_FULL
			oAce.AceType = js_str_isnumber(iAceType) ? iAceType : 0 
			oAce.AceFlags = js_str_isnumber(iAceFlag) ? iAceFlag : 3; // 0 if a file
			oDAcl.AddAce(oAce);
			oAclSD.DiscretionaryAcl = io_acl_reorderdacl(oDAcl);
			return oADs.SetSecurityDescriptor(oAclSD);
			//return io_acl_securitydacl("setdacl",oDAcl,sFileFolder);
		}
		else if(sOpt.match(/setacentfs/ig)){	
			oAce = new ActiveXObject("AccessControlEntry");
			oAce.Trustee =  sTrustee ? sTrustee : "EVERYONE"
			switch(sPerm.toLowerCase()){
				case "f" : oAce.AccessMask = oAcl.ACE_RIGHT_FULL; break;
					case "c" : oAce.AccessMask = oAcl.ACE_RIGHT_READ | oAcl.ACE_RIGHT_WRITE | oAcl.ACE_RIGHT_EXECUTE | oAcl.ACE_RIGHT_DELETE; break;
					case "rx" : oAce.AccessMask = oAcl.ACE_RIGHT_READ | oAcl.ACE_RIGHT_EXECUTE; break;
					case "rw" : oAce.AccessMask = oAcl.ACE_RIGHT_READ | oAcl.ACE_RIGHT_WRITE | oAcl.ACE_RIGHT_DELETE; break;
					case "r" : oAce.AccessMask = oAcl.ACE_RIGHT_READ; break;
					default : oAce.AccessMask = oAcl.ACE_RIGHT_READ; break;
			}
			oAce.AceType = js_str_isnumber(iAceType) ? iAceType : oAcl.ADS_ACETYPE_ACCESS_ALLOWED 
			oAce.AceFlags = js_str_isnumber(iAceFlag) ? iAceFlag : 0; // 0 if a file
			oDAcl.AddAce(oAce);
			return io_acl_securitydacl("setdacl",oDAcl,sFileFolder);
		}
		else if(sOpt.match(/setdacl/ig)){
			oAclSD.DiscretionaryAcl = io_acl_reorderdacl(oDAcl);
			return oADs.SetSecurityDescriptor(oAclSD);
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		js_str_kill(oADs,oAce)
	}
}

/*
Windows 2000 BUG after setting permissions
http://support.microsoft.com/?id=834721

The permission on <folder) are incorrectly ordered, which may cause some entries to be ineffective.
Press OK to continue and sort the permissions correctly, or Cancel to reset permissions

*/

function io_acl_reorderdacl(oDAcl){
	try{// Always need to re-order permissions: http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B279682
		if(!js.adssecurityx) return false
		// Comments in the subroutine explain how the ACL should be ordered.
		// The IADsAccessControlList::AddAce method makes not attempt to properly order the ACE being added.
		var oDAclNew = new ActiveXObject("AccessControlList")
		var oImpDenyDAcl = new ActiveXObject("AccessControlList")
		var oInheritedDAcl =new  ActiveXObject("AccessControlList")
		var oImpAllowDAcl = new ActiveXObject("AccessControlList")
		var oInhAllowDAcl = new ActiveXObject("AccessControlList")
		var oImpDenyObjectDAcl = new ActiveXObject("AccessControlList")
		var oImpAllowObjectDAcl = new ActiveXObject("AccessControlList")
		// Sift the DACL into 5 bins:
		// Inherited Aces
		// Implicit Deny Aces
		// Implicit Deny Object Aces
		// Implicit Allow Aces
		// Implicit Allow object aces
		for(var oItem, oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item(); //  Sort the original ACEs into their appropriate ACLs
			if((oItem.AceFlags && oAcl.ADS_ACEFLAG_INHERITED_ACE) == oAcl.ADS_ACEFLAG_INHERITED_ACE){
				//  Don// t really care about the order of inherited aces.  Since we are adding them to the top of a new list, when they are added back
				//  to the Dacl for the object, they will be in the same order as they were originally.  Just a positive side affect of adding items
				//  of a LIFO ( Last In First Out) type list.
				oInheritedDAcl.AddAce(oItem)
			}
			else {//  We have an Implicit ACE, lets put it the proper pool
				if(oItem.AceType ==  oAcl.ADS_ACETYPE_ACCESS_ALLOWED) oImpAllowDAcl.AddAce(oItem) //  We have an implicit allow ace
				else if(oItem.AceType ==  oAcl.ADS_ACETYPE_ACCESS_DENIED) oImpDenyDAcl.AddAce(oItem) // We have a implicit Deny ACE
				else if(oItem.AceType ==  oAcl.ADS_ACETYPE_ACCESS_ALLOWED_OBJECT) oImpAllowObjectDAcl.AddAce(oItem) //   We have an object allowed ace Does it apply to a property? or an Object?
				else if(oItem.AceType ==  oAcl.ADS_ACETYPE_ACCESS_DENIED_OBJECT) oImpDenyObjectDAcl.AddAce(oItem) //  We have a object Deny ace
				else;
   			}
   		}
		//  Combine the ACEs in the proper order
		
		//  Implicit Deny
		for(var oItem, oEnum = new Enumerator(oImpDenyDAcl); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			oDAclNew.AddAce(oItem)
		}
		//  Implicit Deny Object
		for(var oItem, oEnum = new Enumerator(oImpDenyObjectDAcl); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			oDAclNew.AddAce(oItem)
		}
		//  Implicit Allow
		for(var oItem, oEnum = new Enumerator(oImpAllowDAcl); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			oDAclNew.AddAce(oItem)
		}
		//  Implicit Allow Object
		for(var oItem, oEnum = new Enumerator(oImpAllowObjectDAcl); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			oDAclNew.AddAce(oItem)
		}
		//  Inherited Aces
		for(var oItem, oEnum = new Enumerator(oInheritedDAcl); !oEnum.atEnd(); oEnum.moveNext()){
			oItem = oEnum.item();
			oDAclNew.AddAce(oItem)
		}
		oInheritedDAcl = oImpAllowDAcl = oImpDenyObjectDAcl = oImpDenyDAcl = null
		oDAclNew.AclRevision = oDAcl.AclRevision //  var the appropriate revision level for the DACL
	
		// Replace the Security Descriptor
		var oDAcl = null
		var oDAcl = oDAclNew
		return oDAcl
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_acl_getpermissions(sPermissions,sRegSplit1,sRegSplit2){
	try{ // sPermission: add:woody:f::add:everyone:r where sRegSplit: ::
		if(!js_str_isdefined(sPermissions,sRegSplit1,sRegSplit2)) return false
		var oRe1 = new RegExp(sRegSplit1,"ig"), oRe2 = new RegExp(sRegSplit2,"ig")
		var p = sPermissions.split(oRe1), aPerms = new Array()
		for(var pp, o, i = 0; i < p.length; i++){
			if(pp = (p[i]).split(oRe2)){
				o = new Object();
				o.type = (pp[0]).toLowerCase() // add or del
				o.name = (pp[1]).toLowerCase()
				o.perm = (pp[2]).toLowerCase() // f or rx or rw or r or c
				aPerms.push(o);
			}
		}
		return aPerms;
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_acl_getrights(sAccessMask){
	try{
		var sRight = sAccessMask
		if(sAccessMask == oAcl.SHARE_RIGHT_FULL) sRight = "FULL [ALL]"
		else if(sAccessMask == oAcl.SHARE_RIGHT_CHANGE) sRight = "CHANGE [RWXD]"
		else if(sAccessMask == oAcl.SHARE_RIGHT_READ) sRight = "READ [RX]"
		else if(sAccessMask == oAcl.SHARE_RIGHT_DENY) sRight = "DENY []"
		return sRight
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_acl_getflags(sAceFlags){
	try{
		var sFlags = ""
		if(sAceFlags & oAcl.ADS_ACEFLAG_UNKNOWN) sFlags = sFlags + "U|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_INHERIT_ACE) sFlags = sFlags + "IA1|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE) sFlags = sFlags + "IANP|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_INHERIT_ONLY_ACE) sFlags = sFlags + "IOA|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_INHERITED_ACE) sFlags = sFlags + "IA2|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_INHERIT_FLAGS) sFlags = sFlags + "IF|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_SUCCESSFUL_ACCESS) sFlags = sFlags + "SA|"
		if(sAceFlags & oAcl.ADS_ACEFLAG_FAILED_ACCESS) sFlags = sFlags + "FA|"
		
		return sFlags.substring(0,sFlags.length-1)
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
}

function io_acl_share(sOpt,sComputer,sShareName,sLocSharePath,sDescription,sTrustee,sPerm){
	var oLan
	try{ // http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_securitydescriptor.asp
		if(!js_sys_ping(sComputer)) return false
		oLan = GetObject("WinNT://" + sComputer + "/LanmanServer")
		if(sOpt == "set"){
			var oShare = oLan.Create("FileShare",sShareName)
			oShare.Path = sLocSharePath
			oShare.MaxUserCount = -1 // Unlimeted: Note that workstations has max 10
			oShare.Description = sDescription ? sDescription : ""
			oShare.SetInfo();
			sTrustee = sTrustee ? sTrustee : "EVERYONE"
			sPerm = sPerm ? sPerm : oAcl.SHARE_RIGHT_FULL
			sUNCShare = "\\\\" + sComputer + "\\" + sShareName
			return io_acl_securitydacl("setaceshare",null,sUNCShare,sTrustee,sPerm)			
		}
		else if(sOpt == "get"){
			var oShare = false
			for(var oItem, oEnum = new Enumerator(oLan); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				if((oItem.Name).match(sShareName,"i")){
					oShare = new Object()
					oShare.Name = oItem.Name
					oShare.MaxUserCount = oItem.MaxUserCount
					oShare.AdsPath = oItem.AdsPath
					oShare.Path = oItem.Path
					oShare.Class = oItem.Class
					oShare.CurrentUserCount = oItem.CurrentUserCount
					oShare.description = oItem.description
					oShare.GUID = oItem.GUID
					oShare.HostComputer = oItem.HostComputer
					oShare.Parent = oItem.Parent
					oShare.Schema = oItem.Schema
					break;
				}
			}
			return oShare
		}
		else if(sOpt == "list"){
			if(oo = io_acl_share("get",sComputer,sShareName)){
				s = ""
				for(o in oo) s = s + "\n" + o + " = " + oo[o]
				return s
			}
		}
		else if(sOpt == "del"){
			for(var oItem, oEnum = new Enumerator(oLan); !oEnum.atEnd(); oEnum.moveNext()){
				if((oItem = oEnum.item()) != null){					
					if((oItem.Name).match(sShareName,"i")){
						oLan.Delete("fileshare",sShareName);
						break;
					}
				}
			}
		}
		return false
	}
	catch(e){
		js_log_error(2,e);
		return false;
	}
	finally{
		oLan = null
	}
}