Option Compare Database
Option Explicit

Function cleanDatabase() As Boolean

cleanDatabase = False

Dim errorMsg As New Collection, fso As Object
If Nz(Form__MAIN.releaseNotes, "") = "" Then errorMsg.Add "Empty release notes"

If errorMsg.Count > 0 Then
    Dim msgContents As String, ITEM
    msgContents = ""
    For Each ITEM In errorMsg
        msgContents = msgContents & vbNewLine & ITEM
    Next ITEM
    MsgBox msgContents, vbInformation, "Please fix these issues: "
    GoTo exitThis
End If

If MsgBox("Are you sure? ", vbYesNo, "Just Checking") = vbNo Then GoTo exitThis

addNote "--- STARTING " & Form__MAIN.cmdRepo.Column(2) & " CLEANING PROCEDURE ---"

'---Setup Variables
addNote "establishing variables..."

Set fso = CreateObject("Scripting.FileSystemObject")

TempVars.Add "releaseNum", Form__MAIN.releaseNum.value
TempVars.Add "releaseNotes", Replace(Form__MAIN.releaseNotes.value, "'", "''")
TempVars.Add "responsiblePerson", Form__MAIN.responsiblePerson.value
TempVars.Add "userEmail", Form__MAIN.userEmail.value
TempVars.Add "databaseName", Form__MAIN.cmdRepo.Column(2)

Dim repoLoc As String
repoLoc = Form__MAIN.cmdRepo
TempVars.Add "devFile", getDB

Dim devBackup, devTemp, feFile

'---Open DEV to finalize---
'-if Front End
If Form__MAIN.cmdRepo.Column(2) = "WorkingDB_FE.accdb" Then
    addNote "Opening database for cleaning/compiling"
    Dim dbInput, dbInputRS As Database
    
    Set dbInputRS = OpenDatabase(TempVars!devFile)
    dbInputRS.Execute "DELETE FROM tblPLM"
    dbInputRS.Execute "Delete * from tblSessionVariables"
    dbInputRS.Execute "Update [tblDBinfo] SET [Release] = '" & TempVars!releaseNum & "' WHERE [ID] = 1"
    dbInputRS.Close
    Set dbInputRS = Nothing
    
    Set dbInput = CreateObject("Access.Application")
    dbInput.OpenCurrentDatabase TempVars!devFile
    dbInput.runCommand acCmdCloseAll
    
    Dim checkThis
    Do
        checkThis = dbInput.Run("readyForPublish")
    Loop Until checkThis = True
    
    dbInput.CloseCurrentDatabase
    dbInput.Quit
    
    Dim BEbackup
    TempVars.Add "dbLoc", "\\data\mdbdata\WorkingDB\"
    BEbackup = TempVars!dbLoc & "_backups\prod-BE\"
    
    addNote "Backup backends"
    Call fso.CopyFile(TempVars!dbLoc & "prod-BE\WorkingDB_BE.accdb", BEbackup & TempVars!releaseNum & "_WorkingDB_BE.accdb")
    Call fso.CopyFile(TempVars!dbLoc & "prod-BE\WorkingDB_BE_ChangePointE.accdb", BEbackup & TempVars!releaseNum & "_WorkingDB_BE_ChangePointE.accdb")
    Call fso.CopyFile(TempVars!dbLoc & "prod-BE\WorkingDB_BE_DesignE.accdb", BEbackup & TempVars!releaseNum & "_WorkingDB_BE_DesignE.accdb")
    Call fso.CopyFile(TempVars!dbLoc & "prod-BE\WorkingDB_BE_ProjectE.accdb", BEbackup & TempVars!releaseNum & "_WorkingDB_BE_ProjectE.accdb")
    Call fso.CopyFile(TempVars!dbLoc & "prod-BE\WorkingDB_BE_Sales.accdb", BEbackup & TempVars!releaseNum & "_WorkingDB_BE_Sales.accdb")
End If

addNote "Enable Shift Bypass"

'---Enable Shift---
Call shiftKeyBypass(TempVars!devFile, True)

'---Decompile---
addNote "Decompile"
MsgBox "Hold shift as you click OK - then close the database", vbInformation, "Up Next"
openPath (repoLoc & "decompile.cmd")
MsgBox "Once the database is closed, then click OK", vbInformation, "Up Next"

'---Compact / Repair Dev into Temp---
addNote "Compacting dev into temp file"

TempVars.Add "devTemp", repoLoc & "temp.accdb"
Application.compactRepair TempVars!devFile, TempVars!devTemp
fso.DeleteFile (TempVars!devFile)

'---Compile Temp File---
addNote "Compile"
Dim dbTemp
Set dbTemp = CreateObject("Access.Application")
MsgBox "Hold shift as you click OK", vbInformation, "Up Next"
dbTemp.OpenCurrentDatabase TempVars!devTemp
dbTemp.Visible = False
dbTemp.runCommand acCmdCloseAll

Dim compileMe
Set compileMe = dbTemp.VBE.CommandBars.FindControl(msoControlButton, 578)
If compileMe.Enabled Then compileMe.Execute

dbTemp.runCommand acCmdCompileAndSaveAllModules
dbTemp.CurrentDb.Properties("AllowByPassKey") = True
If fso.FolderExists("H:\wdbBackups\") = False Then MkDir ("H:\wdbBackups\")
addNote "Backup temp into homedrive"
devBackup = "H:\wdbBackups\WorkingDB_Dev_backup.accdb"
Call fso.CopyFile(TempVars!devTemp, devBackup)

addNote "Disable shift bypass"
dbTemp.CurrentDb.Properties("AllowByPassKey") = False
addNote "Close temp file"
dbTemp.CloseCurrentDatabase
dbTemp.Quit
DoEvents

'---Compact Temp into Dev---
addNote "Compacting temp file back into FE"
Application.compactRepair TempVars!devTemp, TempVars!devFile
fso.DeleteFile (TempVars!devTemp)

addNote Form__MAIN.cmdRepo.Column(2) & " CLEANED"

cleanDatabase = True
exitThis:

End Function

Function getRepoInfo(repoLocation) As Boolean
getRepoInfo = False

If Form__MAIN.trackRevisions Then
    'grab lastest revision
    addNote "Getting " & Form__MAIN.cmdRepo & " latest revision"
    
    Dim maxRel
    maxRel = DMax("ID", Form__MAIN.revisionTableName, "databaseName = '" & Form__MAIN.cmdRepo.Column(2) & "'")
    
    Form__MAIN.releaseNum = DLookup("DatabaseVersion", Form__MAIN.revisionTableName, "ID = " & Nz(maxRel, 0))
End If

'find current branch
Form__MAIN.gitbranch = runGitCmd("git branch --show-current")

'list branches
Form__MAIN.gitbranch.RowSource = Replace(Replace(runGitCmd("git branch"), vbLf, ";"), "*", "")
Form__MAIN.gitBranchSelect.RowSource = Form__MAIN.gitbranch.RowSource

Form__MAIN.publishChanges.Visible = Nz(Form__MAIN.cmdRepo.Column(1), "") <> ""

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tblLastUsed SET repoLocation = '" & repoLocation & "' WHERE recordId = 1"
DoCmd.SetWarnings True

getRepoInfo = True
End Function

Function getDB() As String
getDB = Form__MAIN.cmdRepo & Form__MAIN.cmdRepo.Column(2)
End Function

Function shiftKeyBypass(location As String, toggle As Boolean) As Boolean
shiftKeyBypass = False
On Error GoTo errEnableShift

'initialize variables
Dim db As DAO.Database, acc
Dim prop As DAO.Property
Const conPropNotFound = 3270
  
'open the database as an Access object
Set acc = CreateObject("Access.Application")

'open the "database" now within that object
Set db = acc.DBEngine.OpenDatabase(location, False, False)

'run the command
db.Properties("AllowByPassKey") = toggle

GoTo exitThis

errEnableShift:
If Err = conPropNotFound Then
    Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, toggle)
    db.Properties.Append prop
    Resume Next
    GoTo exitThis
End If

MsgBox "Done!"

exitThis: 'clear your objects/detach from the database
db.Close
Set db = Nothing
Set acc = Nothing

shiftKeyBypass = True
End Function

Function runGitCmd(inputCmd As String, Optional dir As String = "current", Optional printAll As Boolean = True, Optional printNone As Boolean = False) As String

Dim wsShell As Object
Dim execObject As Object
Dim sOutput As String
Dim sWorkingDirectory As String

' Set the working directory to your Git repository
If dir = "current" Then
    sWorkingDirectory = Form__MAIN.cmdRepo
Else
    sWorkingDirectory = dir
End If


Set wsShell = CreateObject("WScript.Shell")
wsShell.CurrentDirectory = sWorkingDirectory

Select Case inputCmd
    Case "git commit -a"
        inputCmd = inputCmd & " -m """ & Form__MAIN.releaseNotes & """"
End Select

With CreateObject("WScript.Shell")
    .Run "cmd /c " & inputCmd & " > %temp%\tempgitoutput.txt", 0, True
End With

On Error Resume Next
Dim strOutput
With CreateObject("Scripting.FileSystemObject")
    strOutput = .OpenTextFile(Environ("temp") & "\tempgitoutput.txt").ReadAll()
    .DeleteFile Environ("temp") & "\tempgitoutput.txt"
End With
On Error GoTo 0


Dim arr() As String
arr = Split(strOutput, vbLf)

Dim ITEM, startChanges As Boolean
startChanges = False

If printNone Then GoTo noPrint

DoCmd.SetWarnings False
For Each ITEM In arr
    If ITEM = "" Then startChanges = True
    
    If startChanges = True And printAll = False Then GoTo noPrint
    DoCmd.RunSQL "INSERT INTO tblReleaseTracking(task) VALUES('" & StrQuoteReplace(ITEM) & " ')"
    
Next ITEM

noPrint:
DoCmd.SetWarnings True

On Error Resume Next
Form_sfrmTracking.Requery

moveTrackingToLastRecord

runGitCmd = strOutput

Set execObject = Nothing
Set wsShell = Nothing

End Function

Function moveTrackingToLastRecord()

Dim rs As DAO.Recordset
Dim lNumRec As Long
Dim lNoRecOnForm As Long

Set rs = Form_sfrmTracking.RecordsetClone ' Create a clone of the form's recordset
rs.MoveLast ' Move to the last record in the recordset
lNumRec = rs.RecordCount ' Get the total number of records

' Calculate how many records are visible on the form
lNoRecOnForm = Int(Form_sfrmTracking.InsideHeight / Form_sfrmTracking.Section(acDetail).Height)

' Move the recordset to position the last visible record at the bottom
If lNumRec > lNoRecOnForm Then
    rs.MoveFirst
    rs.Move (lNumRec - lNoRecOnForm)
Else
    rs.MoveFirst ' If fewer records than can be displayed, go to the first
End If

Form_sfrmTracking.Bookmark = rs.Bookmark ' Set the form's bookmark to the calculated position
Form_sfrmTracking.Refresh

Set rs = Nothing ' Release the recordset object

End Function

Function recomposeAccdb(importTo As String)

'---RECOMPOSE---
Dim myComponent
Dim sModuleType
Dim sTempname
Dim sOutstring

Dim myPath, repo
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

importTo = Form__MAIN.cmdRepo & "temp.accdb"
fso.CopyFile Form__MAIN.cmdRepo & Form__MAIN.cmdRepo.Column(2), importTo
repo = "\\data\mdbdata\WorkingDB\build\Send\"
myPath = fso.getparentfoldername(importTo)

addNote "starting Access..."
Dim oApplication
Set oApplication = CreateObject("Access.Application")
addNote "opening " & importTo & " ..."
oApplication.OpenCurrentDatabase importTo
oApplication.runCommand acCmdCloseAll
oApplication.CurrentDb.Properties("AllowByPassKey") = True

Dim folder
Set folder = fso.getfolder(repo)

Dim myFile, objectname, objecttype
For Each myFile In folder.Files
    objecttype = fso.GetExtensionName(myFile.Name)
    objectname = fso.GetBaseName(myFile.Name)
    addNote "Loading " & objectname & " (" & objecttype & ")"

    Select Case objecttype
        Case "form"
        oApplication.LoadFromText acForm, objectname, myFile.Path
        addNote objectname & " LOADED"
        Case "bas"
        oApplication.LoadFromText acModule, objectname, myFile.Path
        addNote objectname & " LOADED"
        Case "mod"
        oApplication.LoadFromText acMacro, objectname, myFile.Path
        addNote objectname & " LOADED"
        Case "rpt"
        oApplication.LoadFromText acReport, objectname, myFile.Path
        addNote objectname & " LOADED"
        Case "qry"
        oApplication.LoadFromText acQuery, objectname, myFile.Path
        addNote objectname & " LOADED"
    End Select
Next

oApplication.runCommand acCmdCompileAndSaveAllModules
oApplication.CloseCurrentDatabase
oApplication.Quit

addNote "Files Imported"

End Function

Function decomposeAccdb(sADPFilename As String, sExportPath As String)

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim myType, myName, myPath, sStubADPFilename As String
myType = fso.GetExtensionName(sADPFilename)
myName = fso.GetBaseName(sADPFilename)
myPath = fso.getparentfoldername(sADPFilename)

sStubADPFilename = Environ("temp") & "\" & myName & "_stub." & myType
addNote sStubADPFilename
addNote "copy stub to " & sStubADPFilename & "..."
fso.CopyFile sADPFilename, sStubADPFilename

addNote "starting Access..."

Dim dbT, accT
Set accT = CreateObject("Access.Application")
Set dbT = accT.DBEngine.OpenDatabase(sStubADPFilename, False, False)

dbT.Properties("AllowByPassKey") = True
dbT.Close
Set dbT = Nothing
accT.Quit
Set accT = Nothing

Dim oApplication
Set oApplication = CreateObject("Access.Application")
addNote "opening " & sStubADPFilename & " ..."
oApplication.OpenCurrentDatabase sStubADPFilename
oApplication.Visible = False

addNote "exporting..."
Dim myObj
Dim delFold
Dim delFile

'delete all files
addNote "  --Clearing Forms Tracking Files---"
If fso.FolderExists(sExportPath & "\Forms\") Then
    Set delFold = fso.getfolder(sExportPath & "\Forms\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing SubForms Tracking Files---"
If fso.FolderExists(sExportPath & "\Forms\SubForms\") Then
    Set delFold = fso.getfolder(sExportPath & "\Forms\SubForms\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing Modules Tracking Files---"
If fso.FolderExists(sExportPath & "\Modules\") Then
    Set delFold = fso.getfolder(sExportPath & "\Modules\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing Macros Tracking Files---"
If fso.FolderExists(sExportPath & "\Macros\") Then
    Set delFold = fso.getfolder(sExportPath & "\Macros\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing Reports Tracking Files---"
If fso.FolderExists(sExportPath & "\Reports\") Then
    Set delFold = fso.getfolder(sExportPath & "\Reports\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing SubReports Tracking Files---"
If fso.FolderExists(sExportPath & "\Reports\SubReports\") Then
    Set delFold = fso.getfolder(sExportPath & "\Reports\SubReports\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing Queries Tracking Files---"
If fso.FolderExists(sExportPath & "\Queries\") Then
    Set delFold = fso.getfolder(sExportPath & "\Queries\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing SubQueries Tracking Files---"
If fso.FolderExists(sExportPath & "\Queries\SubQueries\") Then
    Set delFold = fso.getfolder(sExportPath & "\Queries\SubQueries\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing Tables Tracking Files---"
If fso.FolderExists(sExportPath & "\Tables\") Then
    Set delFold = fso.getfolder(sExportPath & "\Tables\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Clearing VBProject Tracking Files---"
If fso.FolderExists(sExportPath & "\VBProject\") Then
    Set delFold = fso.getfolder(sExportPath & "\VBProject\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

Set delFile = Nothing
Set delFold = Nothing

'---FORMS---
For Each myObj In oApplication.CurrentProject.AllForms
    If Not fso.FolderExists(sExportPath & "\Forms\") Then MkDir (sExportPath & "\Forms\")
    addNote "  exporting form: " & myObj.FullName
    'move all new files
    If Left(myObj.FullName, 1) = "s" Then
        If Not fso.FolderExists(sExportPath & "\Forms\SubForms\") Then MkDir (sExportPath & "\Forms\SubForms\")
        oApplication.SaveAsText acForm, myObj.FullName, sExportPath & "\Forms\SubForms\" & myObj.FullName & ".form"
        splitFormFile (sExportPath & "\Forms\SubForms\" & myObj.FullName & ".form")
    Else
        oApplication.SaveAsText acForm, myObj.FullName, sExportPath & "\Forms\" & myObj.FullName & ".form"
        splitFormFile (sExportPath & "\Forms\" & myObj.FullName & ".form")
    End If
Next

'---MODULES---
For Each myObj In oApplication.CurrentProject.AllModules
    If Not fso.FolderExists(sExportPath & "\Modules\") Then MkDir (sExportPath & "\Modules\")
    addNote "  exporting module: " & myObj.FullName
    oApplication.SaveAsText acModule, myObj.FullName, sExportPath & "\Modules\" & myObj.FullName & ".bas"
Next

For Each myObj In oApplication.CurrentProject.AllMacros
    If Not fso.FolderExists(sExportPath & "\Macros\") Then MkDir (sExportPath & "\Macros\")
    addNote "  exporting macro: " & myObj.FullName
    oApplication.SaveAsText acMacro, myObj.FullName, sExportPath & "\Macros\" & myObj.FullName & ".mod"
Next

'---REPORTS---
For Each myObj In oApplication.CurrentProject.AllReports
    If Not fso.FolderExists(sExportPath & "\Reports\") Then MkDir (sExportPath & "\Reports\")
    addNote "  exporting report: " & myObj.FullName
    If Left(myObj.FullName, 1) = "s" Then
        If Not fso.FolderExists(sExportPath & "\Reports\SubReports\") Then MkDir (sExportPath & "\Reports\SubReports\")
        oApplication.SaveAsText acReport, myObj.FullName, sExportPath & "\Reports\SubReports\" & myObj.FullName & ".rpt"
    Else
        oApplication.SaveAsText acReport, myObj.FullName, sExportPath & "\Reports\" & myObj.FullName & ".rpt"
    End If
Next

'---QUERIES---
For Each myObj In oApplication.CurrentDb.QueryDefs
    If Not Left(myObj.Name, 3) = "~sq" Then 'exclude queries defined by the forms. Already included in the form itself
        If Not fso.FolderExists(sExportPath & "\Queries\") Then MkDir (sExportPath & "\Queries\")
        addNote "  exporting query: " & myObj.Name
        If Left(myObj.Name, 1) = "s" Then
            If Not fso.FolderExists(sExportPath & "\Queries\SubQueries\") Then MkDir (sExportPath & "\Queries\SubQueries\")
            Call writeToTextFile(sExportPath & "\Queries\SubQueries\" & myObj.Name & ".sql", myObj.SQL)
            'oApplication.SaveAsText acQuery, myObj.Name, sExportPath & "\Queries\SubQueries\" & myObj.Name & ".qry"
        Else
            Call writeToTextFile(sExportPath & "\Queries\" & myObj.Name & ".sql", myObj.SQL)
            'oApplication.SaveAsText acQuery, myObj.Name, sExportPath & "\Queries\" & myObj.Name & ".qry"
        End If
    End If
Next

'---TABLES---
For Each myObj In oApplication.CurrentDb.TableDefs
    If Not fso.FolderExists(sExportPath & "\Tables\") Then MkDir (sExportPath & "\Tables\")

    If myObj.Connect = "" Then 'for local tables only, include data
        addNote "  exporting table definition: " & myObj.Name
        oApplication.ExportXML acTable, myObj.Name, sExportPath & "\Tables\" & myObj.Name & "_rows.xml", sExportPath & "\Tables\" & myObj.Name & "_def.xml", , , , acExportAllTableAndFieldProperties
    End If
Next

'---VB PROJECT INFORMATION---
Dim body As String, dictSubValues As Object, dictBody As Object
Set dictSubValues = CreateObject("Scripting.Dictionary")
Set dictBody = CreateObject("Scripting.Dictionary")

addNote "  exporting vbproject information"

For Each myObj In oApplication.VBE.ActiveVBProject.References
    If Not fso.FolderExists(sExportPath & "\VBproject\") Then MkDir (sExportPath & "\VBproject\")
    addNote "  " & myObj.Name & myObj.major & "." & myObj.minor
    dictSubValues.Add myObj.Name & " " & myObj.major & "." & myObj.minor, myObj.FullPath
Next

dictBody.Add "project-name", oApplication.VBE.ActiveVBProject.Name
dictBody.Add "vb-references", dictSubValues

Call writeToTextFile(sExportPath & "\VBproject\VBproject-properties.json", ToJson(dictBody))

'---DB FILE PROPERTIES---
Dim dbtestthis
Set dbtestthis = oApplication.CurrentDb

Set myObj = Nothing
oApplication.CloseCurrentDatabase
oApplication.Quit
Set oApplication = Nothing
Set fso = Nothing

addNote "++ Files Decomposed from " & sADPFilename

End Function

Function writeToTextFile(fileLocation As String, textToWrite As String)

Dim FileNum As Integer

Open fileLocation For Output As #1 ' Open the file for output
Print #1, textToWrite

Close #1

End Function

Function splitFormFile(fileLocation)

Dim FileNum As Integer
Dim DataLine As String
Dim codeLine As Boolean
codeLine = False

Dim myFile As String
myFile = Replace(fileLocation, ".form", ".bas")
Open myFile For Output As #2 ' Open the file for output

FileNum = FreeFile()
Open fileLocation For Input As #1

While Not EOF(FileNum)
    Line Input #1, DataLine ' read in data 1 line at a time
    
    If codeLine Then
        Print #2, DataLine
    End If
    
    If DataLine = "CodeBehindForm" Then codeLine = True
    
Wend

Close #1
Close #2

End Function

Function addNote(noteTxt As String)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO tblReleaseTracking(task) VALUES('" & StrQuoteReplace(noteTxt) & " ')"
DoCmd.SetWarnings True

On Error Resume Next
Form_sfrmTracking.Requery

moveTrackingToLastRecord

DoEvents

End Function