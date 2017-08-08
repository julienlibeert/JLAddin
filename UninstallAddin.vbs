const ADDINFILENAME = "JLAdd-in.xlam"
const UPDATERVERSION = "1.0b"
const FILESTODELETE = 2
dim sAddInDir
dim fso
dim bExcelRunning
dim aFiles(2)
dim i
dim sStatus

afiles(1)="Renamer.xls"
afiles(2)="JLAdd-in.xlam"

sAddinDir =left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
sStatus=""
if msgbox("Are you sure you wish to uninstall the " & ADDINFILENAME & " Add-in and all of it's components?",vbyesno+vbcritical)=vbno then wscript.quit


'close excel
if fnExcelRunning = true then wscript.quit

for i =1 to filestodelete
	if deletefile(saddindir & afiles(i))=true then
		sstatus = sstatus & "File deleted successfully: " & afiles(i) & chr(10)
	else
		sstatus = sstatus & "File could not be deleted: " & afiles(i) & chr(10)
	end if
	if instr(1,afiles(i),".xla")>0 then 'it's an add-in
		If UninstallXLAddin(afiles(i)) = False Then
			sstatus = sstatus & "Registry update failed" & chr(10)
		Else
			sstatus = sstatus & "Registry update successful!" & chr(10)
		end if
	end if
next

msgbox("Uninstallation report:" & chr(10) & sstatus & "Add-in uninstallation completed! Please delete the uninstaller file manually")

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
	    xlversion=gExcel.version
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
                ''''''ease the pressure on the old cpu whilst the user sorts his s#@t out'''''''
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
function DeleteFile(sDir)
Dim MBoxRes 

Deletefile=false
Set fso = CreateObject("Scripting.FileSystemObject")
    If FnFileFolderExists(sDir,true) = True Then
        fso.deletefile sDir
	deletefile=true
    End If
End function
'----------------------------------------------------------------------------------------------
Function UninstallXLAddin(sAddinname) 
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


UninstallXLAddin= False

    While Not bFinish
        If iIndex > 0 Then
            sValueName = sKey & CStr(iIndex)
    
        Else
            sValueName = sKey
        End If
        On Error Resume Next
        sEntry = LCase(WshShell.RegRead(sValueName))    'read out the entry OPEN# (where # is the eventual index)
        If InStr(sEntry, saddinname) > 0 Then
            'The Addin name already exists  delete this entry
            WshShell.Regdelete svaluename
        Else
            iIndex = iIndex + 1
        End If
        sEntry = ""
	if iindex>1000 then	'file does not exist, may already be removed
		uninstallxladdin=true
		exit function
	end if
    Wend

UninstallXLAddin= True

End Function