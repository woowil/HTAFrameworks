'
'' This script contains functions that emulate part of the functionalityof CACLs
'' It requires ADSSECURITY.DLL from the ADSI SDK
''
'' SYNTAX: EDITACL "FILENAME/DIRNAME", "ADD(USER:RIGHT)+DEL(USER:DUMMYVALUE)"
'' FUNCTION: ADDS THE SPECIFIED RIGHTS TO A SINGLE FILE/FOLDER WITHOUT REPLACING
'' ENTIRE ACL i.e. CACLS /E
'' NOTES:
''          Perms should be delimited by a +
''          The Right in DEL(USER:DUMMYVALUE) is required but ignored
''          RIGHT value can be F C or R (Full, Change or Read)
''
'' SYNTAX: REPLACEACL "FILENAME", "ADD(USER:RIGHT)+ADD(USER:RIGHT)"
'' FUNCTION: REPLACES THE EXISTING ACL WITH RIGHTS SPECIFIED i.e. CACLS with no /E switch
'' NOTES:
''          DEL ignored, Still need to specify ADD as code would get to unwieldy otherwise
''
'' SYNTAX: RECURSIVEEDIT "ROOTDIR","ADD(USER:RIGHT)+DEL(USER.RIGHT)"
'' FUNCTION: EDIT ACLS ON ALL DIRECTORIES AND FILE CONTAINED WITHIN ROOTDIR (AND SUBS)
''           i.e. CACLS /E /T
''
'' SYNTAX: RECURSIVEREPLACE "ROOTDIR","ADDUSER(USER:RIGHT)"
'' FUNCTION REPLACE ACLS ON ALL DIRECTORIES AND FILE CONTAINED WITHIN ROOTDIR (AND SUBS)
''          i.e. CACLS /T
''
''Examples
'EditACL "e:\keys\tnk.ico","add(craig:r)+add(domain admins:F)+(Administrators:F)"
'ReplaceACL "e:\keys\tnk.ico","add(everyone:F)+add(domain users:C)"
'RecursiveEdit "e:\keys","add(everyone:F)+del(craig:F)"
'RecursiveReplace "e:\keys","add(everyone:F)"


Function EditACL(filenm, permspart)
     ' Edit permissions on a single file or folder
     set fs=WScript.CreateObject("Scripting.FileSystemObject")
     chkfile=fs.fileexists(filenm) ' make sure the file exists or wscript will crash
     
     if chkfile=true then
          ChangeACLS filenm, permspart, "EDIT", "FILE"
     else
          chkfolder=fs.folderexists(filenm) ' if its not a file, is it a folder ?
          if chkfolder=true Then
               ChangeACLS filenm, permspart, "EDIT", "FOLDER"
          end if
     end if
     
     set fs=nothing
End Function

Function ReplaceACL(filenm, permspart)
'-- Replace ACL on single file or folder-------
     set fs=Wscript.CreateObject("Scripting.FileSystemObject")
     chkfile=fs.fileexists(filenm) ' make sure file exists
     
     if chkfile=true then
          ChangeACLS filenm, permspart, "REPLACE", "FILE"
     else
          chkfolder=fs.folderexists(filenm) ' if its not a file, is it a folder?
          if chkfolder=true then
               ChangeACLS filenm, permspart, "REPLACE", "FOLDER"
          end if
     end if
     
     set fs=nothing
End Function

Function RecursiveEdit(rootfolder,permspart)
'--- Edit ACL's on rootfolder and all its subfolders and files----
     set fs=Wscript.CreateObject("Scripting.FileSystemObject")
     set rfldr=fs.getfolder(rootfolder)
     ChangeACLS rfldr.path, permspart, "EDIT", "FOLDER" 'edit rootfolder first
     
     for each file in rfldr.files
          'edit all files in root folder
          ChangeACLS rfldr.path & "\" & file.name, permspart, "EDIT", "FILE"
     next
     
     for each sfldr in rfldr.subfolders
          RecursiveEdit sfldr, permspart ' recurse through subfolders
     next
     
     set fs=nothing
     set rfldr=nothing
End Function


Function RecursiveReplace(rootfolder,permspart)
'--Replace ACLS on rootfolder and all its subfolders and files ----
     set fs=Wscript.CreateObject("Scripting.FileSystemObject")
     set rfldr=fs.getfolder(rootfolder)
     ChangeACLS rfldr.path, permspart, "REPLACE","FOLDER"
     
     for each file in rfldr.files
          ChangeACLS rfldr.path & "\" & file.name, permspart,"REPLACE","FILE"
     next
     
     for each sfldr in rfldr.subfolders
          RecursiveReplace sfldr, permspart
     next
     
     set fs=nothing
     set rfldr=nothing
End Function


Function ChangeAcls(FILE,PERMS,REDIT,FFOLDER)
'- Edit ACLS of specified file -----
     Const ADS_ACETYPE_ACCESS_ALLOWED = 0
     Const ADS_ACETYPE_ACCESS_DENIED = 1
     Const ADS_ACEFLAG_INHERIT_ACE = 2
     Const ADS_ACEFLAG_SUB_NEW = 9
     
     Set sec = Wscript.CreateObject("ADsSecurity")
     Set sd = sec.GetSecurityDescriptor("FILE://" & FILE)
     Set dacl = sd.DiscretionaryAcl

     'if flagged Replace then remove all existing aces from dacl first
     If ucase(REDIT)="REPLACE" then
          For Each existingAce In dacl
               dacl.removeace existingace
          next
     End if
     
     'break up Perms into individual actions
     cmdArray=split(perms,"+")
   
     for x=0 to ubound(cmdarray)
          tmpVar1=cmdarray(x)
          if ucase(left(tmpVar1,3))="DEL" then
               ACLAction="DEL"
          else
               ACLAction="ADD"
          end if
          
          tmpcmdVar=left(tmpVar1,len(tmpVar1)-1)
          tmpcmdVar=right(tmpcmdVar,len(tmpcmdVar)-4)
          cmdparts=split(tmpcmdVar,":")
          nameVar=cmdparts(0)
          rightVar=cmdparts(1)
          
          ' if flagged edit, delete ACE's belonging to user about to add an ace for
       
          if ucase(REDIT)="EDIT" then
               for each existingAce In dacl
                    trusteeVar=existingAce.trustee
                    if instr(trusteeVar,"\") then
                         trunameVar=right(trusteeVar,len(trusteeVar)-instr(trusteeVar,"\"))
                    else
                         trunameVar=trusteeVar
                    end if
                    
                    uctrunameVar=ucase(trunameVar)
                    ucnameVar=ucase(nameVar)
                    
                    if uctrunameVar=ucnameVar then
                         dacl.removeace existingace
                    end if
               next
          end if

          ' if action is to del ace then following clause skips addace
          if ACLAction="ADD" then
               if ucase(FFOLDER)="FOLDER" then
                    ' folders require 2 aces for user (to do with inheritance)
                    addace dacl, namevar, rightvar, ADS_ACETYPE_ACCESS_ALLOWED, ADS_ACEFLAG_SUB_NEW
                    addace dacl, namevar, rightvar, ADS_ACETYPE_ACCESS_ALLOWED, ADS_ACEFLAG_INHERIT_ACE
               else
                    addace dacl, namevar, rightvar, ADS_ACETYPE_ACCESS_ALLOWED,0
               end if
          end if
     next
     
     for each ace in dacl
     ' for some reason if ace includes "NT AUTHORITY" then existing ace does not get readded to dacl
     
          if instr(ucase(ace.trustee),"NT AUTHORITY\") then
               newtrustee=right(ace.trustee, len(ace.trustee)-instr(ace.trustee, "\"))
               ace.trustee=newtrustee
          end if
     next
     
     ' final sets and cleanup
     sd.DiscretionaryAcl = dacl 
     sec.SetSecurityDescriptor sd

     set sd=nothing
     set dacl=nothing
     set sec=nothing
End Function

Function addace(dacl,trustee, maskvar, acetype, aceflags)
     ' add ace to the specified dacl
     Const RIGHT_READ = &H80000000
     Const RIGHT_EXECUTE = &H20000000
     Const RIGHT_WRITE = &H40000000
     Const RIGHT_DELETE = &H10000
     Const RIGHT_FULL = &H10000000
     Const RIGHT_CHANGE_PERMS = &H40000
     Const RIGHT_TAKE_OWNERSHIP = &H80000
     
     Set ace = CreateObject("AccessControlEntry")
     ace.Trustee = trustee
     
     select case ucase(MaskVar)
     ' specified rights so far only include FC & R. Could be expanded though
     case "F"
          ace.AccessMask = RIGHT_FULL
     case "C"
          ace.AccessMask = RIGHT_READ or RIGHT_WRITE or RIGHT_EXECUTE or      RIGHT_DELETE
     case "R"
          ace.AccessMask = RIGHT_READ or RIGHT_EXECUTE
     end select
     
     ace.AceType = acetype
     ace.AceFlags = aceflags
     dacl.AddAce ace
     set ace=nothing
End Function




'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------


const ADS_RIGHT_DELETE                 = &h10000
  const ADS_RIGHT_READ_CONTROL           = &h20000
  const ADS_RIGHT_WRITE_DAC              = &h40000
  const ADS_RIGHT_WRITE_OWNER            = &h80000
  const ADS_RIGHT_SYNCHRONIZE            = &h100000
  const ADS_RIGHT_ACCESS_SYSTEM_SECURITY = &h1000000
  const ADS_RIGHT_GENERIC_READ           = &h80000000
  const ADS_RIGHT_GENERIC_WRITE          = &h40000000
  const ADS_RIGHT_GENERIC_EXECUTE        = &h20000000
  const ADS_RIGHT_GENERIC_ALL            = &h10000000
  const ADS_RIGHT_DS_CREATE_CHILD        = &h1
  const ADS_RIGHT_DS_DELETE_CHILD        = &h2
  const ADS_RIGHT_ACTRL_DS_LIST          = &h4
  const ADS_RIGHT_DS_SELF                = &h8
  const ADS_RIGHT_DS_READ_PROP           = &h10
  const ADS_RIGHT_DS_WRITE_PROP          = &h20
  const ADS_RIGHT_DS_DELETE_TREE         = &h40
  const ADS_RIGHT_DS_LIST_OBJECT         = &h80
  const ADS_RIGHT_DS_CONTROL_ACCESS      = &h100
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
' 
' ADS_ACETYPE_ENUM
' Ace Type definitions
'
  const ADS_ACETYPE_ACCESS_ALLOWED           = 0
  const ADS_ACETYPE_ACCESS_DENIED            = &h1
  const ADS_ACETYPE_SYSTEM_AUDIT             = &h2
  const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT    = &h5
  const ADS_ACETYPE_ACCESS_DENIED_OBJECT     = &h6
  const ADS_ACETYPE_SYSTEM_AUDIT_OBJECT      = &h7
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' ADS_ACEFLAGS_ENUM
' Ace Flagcd  Constants
'
  const F_UNKNOWN                  = &h1
  const F_INHERIT_ACE              = &h2
  const F_NO_PROPAGATE_INHERIT_ACE  = &h4
  const F_INHERIT_ONLY_ACE         = &h8
  const F_INHERITED_ACE            = &h10
  const F_INHERIT_FLAGS      = &h1f
  const F_SUCCESSFUL_ACCESS        = &h40
  const F_FAILED_ACCESS            = &h80

dim logfile,fso,rootfolder,sec
set fso = CreateObject("scripting.filesystemobject")
'Set sec = CreateObject("ADsSecurity")

'set logfile = fso.createtextfile("c:\temp\ss.txt",true)

' do a recursive check
'set rootfolder = fso.getfolder("C:\temp\hta")
'CheckDir rootfolder


Sub CheckDir(ByVal AFolder)
    'on error resume next
    Dim MoreFolders, TempFolder
    Set MoreFolders = AFolder.SubFolders
    WScript.Echo AFolder.path
    GetSecurity(AFolder.path)
    'AuditFiles(AFolder)    
    For Each TempFolder In MoreFolders
        CheckDir(TempFolder)
    Next
End Sub

sub AuditFiles(afolder)
    'on error resume next
    Dim AFile,AllFiles
    set AllFiles = afolder.files
    For Each AFile In AllFiles
        wscript.echo AFile.path
        GetSecurity(AFile.path)
    Next
end sub

sub GetSecurity(areaname)
    'on error resume next
    dim filesec,ace,dacl
    set filesec = sec.GetSecuritydescriptor("FILE://" & areaname)

    set dacl = filesec.DiscretionaryAcl

    '-- Show the ACEs in the DACL ----
    For Each ace In dacl
        if ace.AceType = 0 then
            wscript.echo "Ace.Trustee: " & ace.Trustee
            wscript.echo "Ace.AccessMask: " & ace.AccessMask & " - " & reportRights(ace.AccessMask )
            wscript.echo "Ace.AceFlags: " & ace.AceFlags & " - " & reportFlags(ace.AceFlags)
            wscript.echo "Ace.AceType: " & ace.AceType
            wscript.echo vbcrlf
            logfile.writeline(areaname & "," & ace.Trustee & "," & reportRights(ace.AccessMask))
        else
            wscript.echo "No access"
        end if
    Next
end sub

function reportRights(val)
    'on error resume next
    Dim s
    ' reports some simple known perms
    if val = 2032127 then
        s = "FULL CONTROL"
    elseif val = 1245631 then
        s = "CHANGE"
    elseif val = 1179817 then
        s= "READ"
    elseif val = 131241 then
        s = "DENY"
    else
        s=val
    end if
    reportRights = s
end Function

function reportFlags(val)
    'on error resume next
    dim s
    if val and F_UNKNOWN then
        s = s & "U|"
    end if
    if val and F_INHERIT_ACE then
        s = s & "IA1|"
    end if
    if val and F_NO_PROPAGATE_INHERIT_ACE then
        s = s & "IANP|"
    end if
    if val and F_INHERIT_ONLY_ACE then
        s = s & "IOA|"
    end if
    if val and F_INHERITED_ACE then
        s = s & "IA2|"
    end if
    if val and F_INHERIT_FLAGS then
        s = s & "IF|"
    end if
    if val and F_SUCCESSFUL_ACCESS then
        s = s & "SA|"
    end if
    if val and F_FAILED_ACCESS then
        s = s & "FA|"
    end if
    reportFlags = s
End Function