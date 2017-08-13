// nOsliw HUI - HTML/HTA Application Framework Library (https://github.com/woowil/HTAFrameworks)
// Copyright (c) 2003-2013, nOsliw Solutions,  All Rights Reserved.

//**Start Encode**


__H.include("HUI@IO@File.js")

var ObjectACL = { //http://www.serverwatch.com/tutorials/article.php/1476871

	ADS_ACEFLAG_SUB_NEW	: 9,

	// Share right
	SHARE_RIGHT_READ			: 1179817,
	SHARE_RIGHT_FULL			: 2032127,
	SHARE_RIGHT_CHANGE			: 1245631,
	SHARE_RIGHT_DENY			: 131241,

	// ACE to specified DACL
	ACE_RIGHT_READ			: 0x80000000,
	ACE_RIGHT_EXECUTE		: 0x20000000,
	ACE_RIGHT_WRITE			: 0x40000000,
	ACE_RIGHT_DELETE		: 0x10000,
	ACE_RIGHT_FULL			: 0x10000000,
	ACE_RIGHT_CHANGE_PERMS		: 0x40000,
	ACE_RIGHT_TAKE_OWNERSHIP        : 0x80000,

	ADS_RIGHT_DELETE                 : 0x10000,
	ADS_RIGHT_READ_CONTROL           : 0x20000,
	ADS_RIGHT_WRITE_DAC              : 0x40000,
	ADS_RIGHT_WRITE_OWNER            : 0x80000,
	ADS_RIGHT_SYNCHRONIZE            : 0x100000,
	ADS_RIGHT_ACCESS_SYSTEM_SECURITY : 0x1000000,
	ADS_RIGHT_GENERIC_READ           : 0x80000000,
	ADS_RIGHT_GENERIC_WRITE          : 0x40000000,
	ADS_RIGHT_GENERIC_EXECUTE        : 0x20000000,
	ADS_RIGHT_GENERIC_ALL            : 0x10000000,
	ADS_RIGHT_DS_CREATE_CHILD        : 0x1,
	ADS_RIGHT_DS_DELETE_CHILD        : 0x2,
	ADS_RIGHT_ACTRL_DS_LIST          : 0x4,
	ADS_RIGHT_DS_SELF                : 0x8,
	ADS_RIGHT_DS_READ_PROP           : 0x10,
	ADS_RIGHT_DS_WRITE_PROP          : 0x20,
	ADS_RIGHT_DS_DELETE_TREE         : 0x40,
	ADS_RIGHT_DS_LIST_OBJECT         : 0x80,
	ADS_RIGHT_DS_CONTROL_ACCESS      : 0x100,

	// Access Control Entry Type Values
 	// Possible values for the IADsAccessContronEntry::AceType property
	ADS_ACETYPE_ACCESS_ALLOWED           : 0,
	ADS_ACETYPE_ACCESS_DENIED            : 0x1,
	ADS_ACETYPE_SYSTEM_AUDIT             : 0x2,
	ADS_ACETYPE_ACCESS_ALLOWED_OBJECT    : 0x5,
	ADS_ACETYPE_ACCESS_DENIED_OBJECT     : 0x6,
	ADS_ACETYPE_SYSTEM_AUDIT_OBJECT      : 0x7,

	// Access Control Entry Inheritance Flags
	// Possible values for the IADsAccessControlEntry::AceFlags property
	ADS_ACEFLAG_UNKNOWN                  : 0x1, //
	ADS_ACEFLAG_INHERIT_ACE              : 0x2, // child objects will inherit ACE of current object
	ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE : 0x4, // prevents ACE inherited by the object from further propagation
	ADS_ACEFLAG_INHERIT_ONLY_ACE         : 0x8, // indicates ACE used only for inheritance (it does not affect permissions on object itself)
	ADS_ACEFLAG_INHERITED_ACE            : 0x10, // indicates that ACE was inherited
	ADS_ACEFLAG_INHERIT_FLAGS            : 0x1F, // indicates that inherit flags are valid (provides confirmation of valid settings)
	ADS_ACEFLAG_SUCCESSFUL_ACCESS        : 0x40, // for auditing success in system audit ACE
	ADS_ACEFLAG_FAILED_ACCESS            : 0x80, // for auditing failure in system audit ACE

	// Registry Permission Type Values
	KEY_QUERY_VALUE 		: 0x0001,
	KEY_SET_VALUE 		: 0x0002,
	KEY_CREATE_SUB_KEY 	: 0x0004,
	KEY_ENUMERATE_SUB_KEYS 	: 0x0008,
	KEY_NOTIFY 		: 0x0010,
	KEY_CREATE_LINK		: 0x0020,
	DELETE 			: 0x00010000,
	READ_CONTROL 		: 0x00020000,
	WRITE_DAC 		: 0x00040000,
	WRITE_OWNER 		: 0x00080000
}

__H.register(__H.IO.File,"ACL","ACL Security",function ACL(){
	// see http://www.visualbasicscript.com/topic.asp?TOPIC_ID=2220 http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B279682
	var TXT_DLL_NOT_EXIST = "Missing argument or unable to load componet ADsSecurity.DLL"
	var TXT_DLL_NOT_LOADED = "Unable to load DLL"
	var oAcl = ObjectACL;
	var s_filefolder = ""
	this.bSetDescriptor = false
	var o_adssecurityx = null
	var s_adsecuritydll = ""
	
	this.initialize = function initialize(sFileADSecurityDLL){
		try{
			
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}
	
	this.setADSecurityDLL = function setADSecurityDLL(sFile){
		try{
			// sFile: The ActiveX Control of ADsSecurity.dll
			if(!o_adssecurityx && !(o_adssecurityx = this.regActiveX("ADsSecurity",sFile))){
				throw new Error(__HLog.errorCode("error"),TXT_DLL_NOT_EXIST)
			}
			s_adsecuritydll = sFile
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}

	}
	
	this.getADSecurityDLL = function getADSecurityDLL(){
		return o_adssecurityx
	}
	
	var setFileFolder = function setFileFolder(sFileFolder){
		try{
			if(!oFso.FolderExists(sFileFolder) && !oFso.FileExists(sFileFolder)){
				throw new Error(__HLog.errorCode("error"),"File/Folder not exist: " + sFileFolder)
			}
			return (s_filefolder == sFileFolder)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setACL = function setACL(sFileFolder,bSetDescriptor){
		try{			
			if(!o_adssecurityx) return false
			if(!setFileFolder(sFileFolder)) return false
			
			this.AclSD = o_adssecurityx.GetSecurityDescriptor("FILE://" + sFileFolder)
			this.DisACL = this.AclSD.DiscretionaryAcl;
			this.Ace = new ActiveXObject("AccessControlEntry");
			this.bSetDescriptor = !!bSetDescriptor
			
			return true
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getACL = function getACL(){
		try{
			return this.DisACL;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.removeACE = function removeACE(){
		try{ // Removes all existing aces from dacl first
			for(var oEnum = new Enumerator(this.DisACL); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				this.DisACL.RemoveAce(oItem);
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setNTAuthority = function setNTAuthority(sFileFolder){
		try{ // For some reason if ace includes "NT AUTHORITY" then existing ace does not get readed to dacl
			if(!this.setACL(sFileFolder)) return false
			var oRe = /nt authority\\(.*)/ig
			for(var oEnum = new Enumerator(this.DisACL); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if((oItem.Trustee).match(oRe)){
					oItem.Trustee = RegExp.$1
				}
			}
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getOwner = function getOwner(){
		try{
			return this.AclSD.Owner;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getACLObject = function getACLObject(){
		try{
			var aAcl = []
			for(var oEnum = new Enumerator(this.DisACL); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(oItem.AceType == 0){
					var o = {
						Trustee    : oItem.Trustee, // BUILTIN\Administrators
						AccessMask : oItem.AccessMask, // 1245631
						Rights     : this.getRights(oItem.AccessMask), // CHANGE or FULL CONTROL or READ or WRITE
						AceFlags   : oItem.AceFlags,
						Flags      : this.getFlags(oItem.AceFlags),
						AceType    : oItem.AceType
					}
					aAcl.push(o);
				}
			}
			return aAcl;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getACLObjectSimple = function getACLObjectSimple(){
		try{
			var aAcl = []
			for(var oEnum = new Enumerator(this.DisACL); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item();
				if(oItem.AceType == 0 && oItem.AccessMask >= 0){ // Ignore duplicates
					var o = {
						Trustee : oItem.Trustee, // BUILTIN\Administrators
						Rights  : this.getRights(oItem.AccessMask) // CHANGE or FULL CONTROL or READ or WRITE
					}
					if(typeof(o.Rights) == "string") aAcl.push(o); // ignore undefined
				}
			}
			return aAcl;
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getACLString = function getACLString(sFileFolder){
		try{
			var s = "Trustee; AccessMask; Rights; AceFlags; Flags; AceType", a
			if(a = this.getACLSObject(sFileFolder)){
				for(var i = 0, iLen = a.length; i < iLen; i++){
					s = s + "\n" + a[i].Trustee + ";" + a[i].AccessMask + ";" + a[i].Rights + ";" + a[i].AceFlags + ";" + a[i].Flags + ";" + a[i].AceType;
				}
			}
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getACLStringSimple = function getACLStringSimple(sFileFolder){
		try{
			var s = "Trustee; Rights", a
			if(a = this.getACLObjectSimple(sFileFolder)){
				for(var i = 0, iLen = a.length; i < iLen; i++){
					s = s + "\n" + a[i].Trustee + ";" + a[i].Rights;
				}
			}
			return s
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setFile = function setFile(sFileFolder,sPermissions,bSetFolder,sOpt3,sOpt4){
		try{
			var sSplit1 = sOpt3 ? sOpt3 : "::"
			var sSplit2 = sOpt4 ? sOpt4 : ":"
			if(__H.isStringEmpty(sPermissions)) return false
			var a = this.getPermissions(sPermissions,sSplit1,sSplit2);
			var oRe = /[a-z ]*\\{0,1}([a-z ]+)/ig
			for(var i = 0, iLen = a.length; i < iLen; i++){
				var oPerm = a[i];
				// delete ACE's belonging to user about to add an ace for
				for(var oEnum = new Enumerator(this.DisACL); !oEnum.atEnd(); oEnum.moveNext()){
					var oItem = oEnum.item();
					var sName = ((oItem.Trustee).match(oRe) ? RegExp.$1 : oItem.Trustee).toLowerCase();
					if(sName == oPerm.name){
						this.DisACL.RemoveAce(oItem)
					}
				}
				// if action is to del ace then following clause skips addace
				if(oPerm.type == "add"){
					var b1, b2
					if(bSetFolder){
						b1 = this.setACENTFS(oPerm.name,oPerm.perm,oAcl.ADS_ACETYPE_ACCESS_ALLOWED,oAcl.ADS_ACEFLAG_SUB_NEW);
						b2 = this.setACENTFS(oPerm.name,oPerm.perm,oAcl.ADS_ACETYPE_ACCESS_ALLOWED,oAcl.ADS_ACEFLAG_INHERIT_ACE);
					}
					else{
						b1 = this.setACENTFS(this.DisACL,oPerm.name,oPerm.perm,oAcl.ADS_ACETYPE_ACCESS_ALLOWED,0)
						b2 = true
					}
					if(!b1 || !b2) sResult = false
				}
			}
			this.setNTAuthority(sFileFolder);
			if(this.bSetDescriptor) this.setDACL(this.DisACL,sFileFolder);
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setFolder = this.setFile;

	this.getRights = function getRights(sAccessMask){
		try{
			var sRight = sAccessMask
			if(sAccessMask == oAcl.SHARE_RIGHT_FULL) sRight = "FULL [ALL]"
			else if(sAccessMask == oAcl.SHARE_RIGHT_CHANGE) sRight = "CHANGE [RWXD]"
			else if(sAccessMask == oAcl.SHARE_RIGHT_READ) sRight = "READ [RX]"
			else if(sAccessMask == oAcl.SHARE_RIGHT_DENY) sRight = "DENY []"
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setACEShare = function setACEShare(oDAcl,sFileFolder,sTrustee,sPerm,iAceType,iAceFlag){
		try{
			if(!this.setACL(sFileFolder)) return false
			this.DisACL = oDAcl ? oDAcl : this.AclSD.DiscretionaryAcl;
			this.Ace.Trustee = sTrustee ? sTrustee : "EVERYONE"
			this.Ace.AccessMask = sPerm ? sPerm : oAcl.SHARE_RIGHT_FULL
			this.Ace.AceType = __H.isNumber(iAceType) ? iAceType : 0
			this.Ace.AceFlags = __H.isNumber(iAceFlag) ? iAceFlag : 3; // 0 if a file
			this.DisACL.AddAce(oAce);
			this.AclSD.DiscretionaryAcl = this.reorderACL(this.DisACL);
			return oADs.SetSecurityDescriptor(this.AclSD);
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setACENTFS = function setACENTFS(sFileFolder,sTrustee,sPerm,iAceType,iAceFlag){
		try{
			this.Ace.Trustee =  !__H.isStringEmpty(sTrustee) ? sTrustee : "EVERYONE"
			switch(sPerm.toLowerCase()){
				case "f" : oAce.AccessMask = oAcl.ACE_RIGHT_FULL; break;
				case "c" : oAce.AccessMask = oAcl.ACE_RIGHT_READ | oAcl.ACE_RIGHT_WRITE | oAcl.ACE_RIGHT_EXECUTE | oAcl.ACE_RIGHT_DELETE; break;
				case "rx" : oAce.AccessMask = oAcl.ACE_RIGHT_READ | oAcl.ACE_RIGHT_EXECUTE; break;
				case "rw" : oAce.AccessMask = oAcl.ACE_RIGHT_READ | oAcl.ACE_RIGHT_WRITE | oAcl.ACE_RIGHT_DELETE; break;
				case "r" : oAce.AccessMask = oAcl.ACE_RIGHT_READ; break;
				default : oAce.AccessMask = oAcl.ACE_RIGHT_READ; break;
			}
			oAce.AceType = __H.isNumber(iAceType) ? iAceType : oAcl.ADS_ACETYPE_ACCESS_ALLOWED
			oAce.AceFlags = __H.isNumber(iAceFlag) ? iAceFlag : 0; // 0 if a file
			this.DisACL.AddAce(oAce);
			return this.setDACL(this.DisACL,sFileFolder);
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.setDACL = function setDACL(oDAcl,sFileFolder){
		try{
			if(!this.setACL(sFileFolder)) return false
			this.DisACL = oDAcl ? oDAcl : this.AclSD.DiscretionaryAcl;
			this.AclSD.DiscretionaryAcl = this.reorderACL(this.DisACL);
			return oADs.SetSecurityDescriptor(this.AclSD);
		}
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.getPermissions = function getPermissions(sPermissions,sRegSplit1,sRegSplit2){
		try{ // sPermission: add:woody:f::add:everyone:r where sRegSplit: ::
			if(__H.isStringEmpty(sPermissions,sRegSplit1,sRegSplit2)) return false
			var oRe1 = new RegExp(sRegSplit1,"ig"), oRe2 = new RegExp(sRegSplit2,"ig")
			var p = sPermissions.split(oRe1), a = []
			for(var pp, i = 0, iLen = p.length; i < iLen; i++){
				if(pp = (p[i]).split(oRe2)){
					var o = {};
					o.type = (pp[0]).toLowerCase() // add or del
					o.name = (pp[1]).toLowerCase()
					o.perm = (pp[2]).toLowerCase() // f or rx or rw or r or c
					a.push(o);
				}
			}
			return a;
		}
		catch(e){
			__HLog.error(ee,this)
			return false;
		}
	}

	this.getFlags = function getFlags(sAceFlags){
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
		catch(ee){
			__HLog.error(ee,this)
			return false
		}
	}

	this.reorderACL = function reorderACL(oDAcl){
		try{
			// Always need to re-order permissions: http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B279682
			// Comments in the subroutine explain how the ACL should be ordered.
			// The IADsAccessControlList::AddAce method makes not attempt to properly order the ACE being added.
			/*
			Windows 2000 BUG after setting permissions
			http://support.microsoft.com/?id=834721__H

			The permission on <folder) are incorrectly ordered, which may cause some entries to be ineffective.
			Press OK to continue and sort the permissions correctly, or Cancel to reset permissions

			*/
			var oDAclNew = new ActiveXObject("AccessControlList")
			var oImpDenyDAcl = new ActiveXObject("AccessControlList")
			var oInheritedDAcl =new  ActiveXObject("AccessControlList")
			var oImpAllowDAcl = new ActiveXObject("AccessControlList")
			var oImpDenyObjectDAcl = new ActiveXObject("AccessControlList")
			var oImpAllowObjectDAcl = new ActiveXObject("AccessControlList")
			// Sift the DACL into 5 bins:
			// Inherited Aces
			// Implicit Deny Aces
			// Implicit Deny Object Aces
			// Implicit Allow Aces
			// Implicit Allow object aces
			for(var oEnum = new Enumerator(oDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				var oItem = oEnum.item(); //  Sort the original ACEs into their appropriate ACLs
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
			for(oEnum = new Enumerator(oImpDenyDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				oDAclNew.AddAce(oItem)
			}
			//  Implicit Deny Object
			for(oEnum = new Enumerator(oImpDenyObjectDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				oDAclNew.AddAce(oItem)
			}
			//  Implicit Allow
			for(oEnum = new Enumerator(oImpAllowDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				oDAclNew.AddAce(oItem)
			}
			//  Implicit Allow Object
			for(oEnum = new Enumerator(oImpAllowObjectDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				oDAclNew.AddAce(oItem)
			}
			//  Inherited Aces
			for(oEnum = new Enumerator(oInheritedDAcl); !oEnum.atEnd(); oEnum.moveNext()){
				oItem = oEnum.item();
				oDAclNew.AddAce(oItem)
			}

			oInheritedDAcl = oImpAllowDAcl = oImpDenyObjectDAcl = oImpDenyDAcl = null
			oDAclNew.AclRevision = oDAcl.AclRevision //  var the appropriate revision level for the DACL

			// Replace the Security Descriptor
			this.DAcl = null
			return (this.DAcl = oDAclNew)
		}
		catch(ee){
			__HLog.error(ee,this)
			return false;
		}
	}
})

var __HACL = new __H.IO.File.ACL()

