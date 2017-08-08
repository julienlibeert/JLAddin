const ADDINFILENAME = "JLAdd-In.xlam"	'this filename will be used to determine default folder for add-in installation
const UPDATERVERSION = "1.0a"	'last update 16/12/2011
const FILESTOMOVE =2
const UNINSTALLERFILE="UninstallAddin"

dim sAddInDest,sDestDir
dim sAddInStart,sStartDir
dim sServVersion
dim sLocVersion
dim sStatus
dim bUpdate
dim aFiles(2)

afiles(1)="Renamer.xls"
afiles(2)="JLAdd-in.xlam"

bUpdate=false

sStatus=""

sstartdir =left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
sAddinstart=sstartdir & ADDINFILENAME

'first check if arguments are included with the directory of the add-in
if wscript.arguments.count > 0 then
	sAddInDest =wscript.arguments(0)	'pushed argument  is "c:\folders\add-in\addin.xlam"
	sdestdir=left( saddindest, len(saddindest)-len(split(saddindest,"\")(ubound(split(saddindest,"\")))))
	bupdate = wscript.arguments(1)	'update?
else
	sAddInDest = fnGetAddInDest	'get the most likely destination for the add-in
	if saddindest="" then
		msgbox("Installation of add-in aborted")
		wscript.quit
	else
		sdestdir=left( saddindest, len(saddindest)-len(split(saddindest,"\")(ubound(split(saddindest,"\")))))
	end if	
end if

'Now, copy the server version to the local destination
if fnExcelRunning = true then wscript.quit
iSuccess=0
for i=1 to FILESTOMOVE
	if movefile(sstartdir & afiles(i), sdestdir & afiles(i))=true then sstatus = sstatus & "Successful transfer of " & afiles(i) & chr(10) else sstatus=sstatus & "Transfer failed for " & afiles(i) & chr(10)
	if instr(1,afiles(i),".xla")>0 then 'it is an add-in
		If AddXLAddin(sdestdir & afiles(i)) = true Then sstatus = sstatus & "Registry update successful (" & afiles(i) & ")" & chr(10) else sstatus =sstatus &  "Registry update failed (" & afiles(i) & ")" & chr(10)
	end if
next

'now copy the uninstaller file
if movefile(sstartdir & uninstallerfile & ".txt", sdestdir & uninstallerfile & ".vbs")=true then
	sstatus = sstatus & "Uninstaller successfully installed" & chr(10)
else
	sstatus = sstatus & "Uninstaller failed to install" & chr(10)
end if

msgbox("Installation report: " & chr(10) & sstatus & "Installation completed")

'----------------------------------------------------------------------------------------------
'**********Functions**********
function fnExcelRunning
strComputer = "."
dim gExcel
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set WSHShell = wscript.CreateObject("wscript.shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

UN = ucase(WshNetwork.UserName) 'Get GID
dim TimeOut
dim Response
TimeOut = 0
Response = 0

Response = msgbox ("You are about to run a macro which will update excel" & vbcrlf & _
    "Excel will be closed (you will be asked to save any unsaved workbooks first" & vbcrlf & _
    "Do you wish to continue?",36, "ATTENTION!")
if Response = 6 then
    dim objProcessList
            ''''''shut down excel'''''''
            on error resume next
            Set gExcel = GetObject(,"Excel.Application")
            gExcel.visible = true
            gExcel.displayalerts = true
            gExcel.activeworkbook.close
            gExcel.application.quit
            gExcel = ""
            on error goto 0

            do   ''''''start a loop to wait for all instances of excel to close before continuing'''''
                ''''''find excel in task manager process list''''''
                Set objWMIService = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                Set colProcessList = objWMIService.ExecQuery _
                ("Select * from Win32_Process Where Name = 'Excel.exe'")
                x = false
                '''''if excel is open change x to true'''''''
                for each objprocess in colprocesslist
                x = true
                'msgbox objprocess.name
                next    
                ''''''if x is false then no excel apps are open so we can carry on with the update''''''''''
                if x = false then 
                exit do
                end if
                ''''''ease the pressure on the old cpu''
                wscript.sleep 2000
                '''prevent an infinite loop incase the user has gone home or cancelled the excel shutdown !!''''
                Timeout = TimeOut + 2000
                'msgbox "timeout = " & timeout
                ''''I give the user 2 minutes to comply, otherwise, terminate process'''''
                if timeout => 180000 then
                msgbox "macro update timed out, please close any open excel spreadsheets and try again"
                wscript.quit
                end if         
            loop
	fnexcelrunning=false
'	msgbox "Excel ready"
else
	fnexcelrunning=true
	msgbox "No update performed"
end if
end function
'----------------------------------------------------------------------------------------------
function fnGetAddInDest() 	'this function gets the addin destination (including file name)
dim DefaultFolderDir
Const MY_COMPUTER = &H11&
Const WINDOW_HANDLE = 0
Const OPTIONS = 0

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_COMPUTER)
Set objFolderItem = objFolder.Self
set oXL = CreateObject("Excel.Application")

defaultfolderdir = fnGetDefaultFolder

if right(defaultfolderdir,1)<>"\" then	'add-in is already installed
	'compare versions:
	oxl.workbooks.open defaultfolderdir
	slocversion = oxl.workbooks(addinfilename).worksheets("Settings").range("Setversion").value
	oxl.workbooks(addinfilename).Close
	oxl.application.quit
	oxl.workbooks.open sAddInStart
	sservversion = oxl.workbooks(addinfilename).worksheets("Settings").range("Setversion").value
	oxl.workbooks(addinfilename).Close
	oxl.application.quit
	if compareversion(slocversion,sservversion) = true then 	'update is newer
		if msgbox("You currently have version " & slocversion & " installed, would you like to install the newer " & sservversion & " version?",vbokcancel)=vbok then
			fngetaddindest=defaultfolderdir
		else
			fngetaddindest=""
		end if
		exit function
	else	'local version is newer
		msgbox "You have the latest version available for this software (" & sservversion & ")",vbokonly
		fngetaddindest=""
		exit function
	end if
end if

if msgbox("It is recommended to install this add-in by updating a previous version of this add-in if available. Are you sure you wish to proceed with this blank installation?", vbyesno) = vbno then
	fngetaddindest =""
	exit function
end if

select case msgbox("Use default folder for add-in installation?" & chr(10) & defaultfolderdir,vbyesnocancel)
	case vbcancel
		fngetaddindest =""
		exit function
	case vbyes
		if fnfilefolderexists(defaultfolderdir,false)=false then call createfolderdir(defaultfolderdir)
		fngetaddindest=defaultfolderdir & ADDINFILENAME
	case vbno	'browse for folder
		strPath = objFolderItem.Path
		Set objShell = CreateObject("Shell.Application")
		Set objFolder = objShell.BrowseForFolder _
			(WINDOW_HANDLE, "Select a folder:", OPTIONS, strPath) 

		If objFolder Is Nothing Then
			fngetaddindest=""
		    	exit function
		End If

		Set objFolderItem = objFolder.Self
		fngetaddindest= objFolderItem.Path & ADDINFILENAME
end select

end function

'----------------------------------------------------------------------------------------------
Function fnGetDefaultFolder()
'this function finds the default folders for add-ins on this computer and checks if they exist
dim defaultfolder

defaultfolder= "C:\Users\" & WScript.CreateObject("WScript.Network").UserName & "\AppData\Roaming\Microsoft"

if fnfilefolderexists(defaultfolder & "\AddIns",false) then
	fngetdefaultfolder=defaultfolder & "\AddIns\" 	'end with backslash to indicate add-in not yet installed
	if fnfilefolderexists(defaultfolder & "\AddIns\" & ADDINFILENAME,true) then 'the actual add-in has already been installed
		fngetdefaultfolder=defaultfolder & "\AddIns\" & ADDINFILENAME
		exit function		
	end if
end if

if fnfilefolderexists(defaultfolder & "\invoegtoepassingen",false) then
	fngetdefaultfolder=defaultfolder & "\invoegtoepassingen\" 'end with backslash to indicate add-in not yet installed
	if fnfilefolderexists(defaultfolder & "\invoegtoepassingen\" & ADDINFILENAME,true) then 'the actual add-in has already been installed
		fngetdefaultfolder=defaultfolder & "\invoegtoepassingen\" & ADDINFILENAME
		exit function		
	end if
else	'no add-in folders exist, then create them!
	fngetdefaultfolder=defaultfolder & "\AddIns\"
end if

end function

'----------------------------------------------------------------------------------------------
Function FnFileFolderExists(sFilePath , bfile )
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If bfile = True Then
    FnFileFolderExists = fso.fileexists(sFilePath)
Else
    FnFileFolderExists = fso.folderexists(sFilePath)
End If

End Function
'----------------------------------------------------------------------------------------------
Sub CreateFolderDir(FolderDir)
Dim fso
dim aFolder
dim fol
dim i

Set fso = CreateObject("Scripting.FileSystemObject")

afolder=split(folderdir,"\")
for i=0 to ubound(afolder)-1
	fol= fol & afolder(i) & "\" 		
	If Not fso.FolderExists(fol) Then fso.CreateFolder (fol)
next

End sub
'----------------------------------------------------------------------------------------------
function MoveFile(sStart, sDest)
'sStart is of the form: C:\file.txt
'sDest is of the form: C:\folder\file.txt
Dim MBoxRes 
dim sDestDir

movefile=false
Set fso = CreateObject("Scripting.FileSystemObject")

sdestdir=left( sdest, len(sdest)-len(split(sdest,"\")(ubound(split(sdest,"\")))))
    If FnFileFolderExists(sDest,true) = True Then
	if bupdate = true then
		mboxres=vbyes
	else
	        MBoxRes = vbyes 'MsgBox(sDest & " already exists, overwrite?", vbYesNoCancel, "Overwrite?")
	end if
        If (MBoxRes = vbNo) Then
           MoveFile = ""
           Exit function
        End If
        If MBoxRes = vbYes Then
            fso.deletefile sDest
        End If
        If MBoxRes = vbCancel Then
            MoveFile = ""
            Exit function
        End If
    elseif fnfilefolderexists(sdestdir,false)=false then
	call createfolderdir(sdestdir)
    End If
    fso.copyFile sStart, sDest   'now the original file is in the targetpath
    movefile=true
End function
'----------------------------------------------------------------------------------------------
Function CompareVersion(byval CurVers, byval ServVers)
'this function will compare two version numbers and return true if an update is available
'version numbers are of following format: 1.1a, 1.1c or 1.5
Dim i
Dim CurVersNum, ServVersNum
Dim CurVersLet, ServVersLet

CompareVersion = False

If CInt(Split(CurVers, ".")(0)) < CInt(Split(ServVers, ".")(0)) Then 'this is a completely new version number
    CompareVersion = True
    Exit Function
Else
    CurVers = Split(CurVers, ".")(1)
    ServVers = Split(ServVers, ".")(1)
End If


CurVersNum = 0: CurVersLet = ""
ServVersNum = 0: ServVersLet = ""

i = 1
Do While i <= Len(CurVers)
    If IsNumeric(Mid(CurVers, i, 1)) = True Then
        CurVersNum = CurVersNum & Mid(CurVers, i, 1)
    Else
        CurVersLet = CurVersLet & UCase(Mid(CurVers, i, 1))
    End If
    i = i + 1
Loop
i = 1
Do While i <= Len(ServVers)
    If IsNumeric(Mid(ServVers, i, 1)) = True Then
        ServVersNum = ServVersNum & Mid(ServVers, i, 1)
    Else
        ServVersLet = ServVersLet & UCase(Mid(ServVers, i, 1))
    End If
    i = i + 1
Loop

If CurVersNum < ServVersNum Then
    CompareVersion = True
ElseIf CurVersNum = ServVersNum Then
    i = 1
    While i <= Len(CurVersLet) And i <= Len(ServVersLet)
        If Asc(Mid(CurVersLet, i, 1)) < Asc(Mid(ServVersLet, i, 1)) Then
            CompareVersion = True
            Exit Function
        ElseIf Asc(Mid(CurVersLet, i, 1)) > Asc(Mid(ServVersLet, i, 1)) Then
            CompareVersion = False
            Exit Function
        End If
        i = i + 1
    Wend
    'if you get this far only the length of the version letter part will determine which is newer
    CompareVersion = CBool(Len(CurVersLet) < Len(ServVersLet))
Else
    CompareVersion = False
End If

End Function

'----------------------------------------------------------------------------------------------
Function AddXLAddin(sAddin) 
sAddin = LCase(sAddin)
Set WshShell = CreateObject("WScript.Shell")
Dim bFinish: bFinish = False
Dim iIndex: iIndex = 0
Dim sEntry: sEntry = ""
Dim sKey: sKey = ""
Dim Quote: Quote = """"

Dim aPath: aPath = Split(sAddin, "\")
Dim sAddinFile: sAddinFile = sAddin 'file including path
Dim sAddinFile1: sAddinFile1 = aPath(UBound(aPath)) 'only filename
Dim sAddinPath: sAddinPath = Left(sAddin, InStr(sAddin, "\" & sAddinFile1)) 'only path
Dim sValueName: sValueName = ""
dim XLVersion
dim XL

set oXL = createobject("Excel.Application")

xlversion = oxl.version

call uninstallxladdin(saddinfile1)	'first delete existing versions from Registry

        sKey = "HKCU\Software\Microsoft\Office\" & xlversion & "\Excel\Options\OPEN"

AddXLAddin = False

If sKey <> "" Then
 
'sAddinFile=left(sAddinFile,Len(sAddinFile)-1)
    While Not bFinish
        If iIndex > 0 Then
            sValueName = sKey & CStr(iIndex)
    
        Else
            sValueName = sKey
        End If
        On Error Resume Next
        sEntry = LCase(WshShell.RegRead(sValueName))    'read out the entry OPEN# (where # is the eventual index)

     If Len(sEntry) < 3 Then
            'No more entries

            WshShell.RegWrite sValueName, Quote & sAddin & Quote, "REG_SZ"
            bFinish = True
        ElseIf InStr(sEntry, ADDINFILENAME) > 0 Then
            'The Addin name already exists  (replace existing)
            WshShell.RegWrite sValueName, Quote & sAddin & Quote, "REG_SZ"
            bFinish = True
        Else
            iIndex = iIndex + 1
        End If
        sEntry = ""
    Wend
    AddXLAddin = True
End If

End Function

'----------------------------------------------------------------------------------------------
sub UninstallXLAddin(sAddinname) 
sAddin = LCase(sAddin)
Set WshShell = CreateObject("WScript.Shell")
Dim bFinish: bFinish = False
Dim iIndex: iIndex = 0
Dim sEntry: sEntry = ""
Dim sKey: sKey = ""

Dim sValueName: sValueName = ""
dim XLVersion
dim XL

set oXL = createobject("Excel.Application")

xlversion = oxl.version

        sKey = "HKCU\Software\Microsoft\Office\" & xlversion & "\Excel\Options\OPEN"


    While Not bFinish
        If iIndex > 0 Then
            sValueName = sKey & CStr(iIndex)
    
        Else
            sValueName = sKey
        End If
        On Error Resume Next
        sEntry = LCase(WshShell.RegRead(sValueName))    'read out the entry OPEN# (where # is the eventual index)
        If InStr(ucase(sEntry), ucase(saddinname)) > 0 Then
            'The Addin name already exists  delete this entry
            WshShell.Regdelete svaluename
        End if
        
	iIndex = iIndex + 1
        sEntry = ""
	if iindex>1000 then	'file does not exist, may already be removed
		exit sub
	end if
    Wend

End Sub