Option Compare Database
Option Explicit

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
db.Properties("AllowByPassKey") = True

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

Function runGitCmd(inputCmd As String, Optional dir As String = "current") As String

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

Dim item
For Each item In arr
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO tblReleaseTracking(task) VALUES('" & StrQuoteReplace(item) & " ')"
    DoCmd.SetWarnings True
Next item

On Error Resume Next
Form_frmTracking.Requery

moveTrackingToLastRecord

runGitCmd = strOutput

Set execObject = Nothing
Set wsShell = Nothing

End Function

Function moveTrackingToLastRecord()

Dim rs As DAO.Recordset
Dim lNumRec As Long
Dim lNoRecOnForm As Long

Set rs = Form_frmTracking.RecordsetClone ' Create a clone of the form's recordset
rs.MoveLast ' Move to the last record in the recordset
lNumRec = rs.RecordCount ' Get the total number of records

' Calculate how many records are visible on the form
lNoRecOnForm = Int(Form_frmTracking.InsideHeight / Form_frmTracking.Section(acDetail).Height)

' Move the recordset to position the last visible record at the bottom
If lNumRec > lNoRecOnForm Then
    rs.MoveFirst
    rs.Move (lNumRec - lNoRecOnForm)
Else
    rs.MoveFirst ' If fewer records than can be displayed, go to the first
End If

Form_frmTracking.Bookmark = rs.Bookmark ' Set the form's bookmark to the calculated position
Form_frmTracking.Refresh

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
addNote "  --Deleting Forms"
If fso.FolderExists(sExportPath & "\Forms\") Then
    Set delFold = fso.getfolder(sExportPath & "\Forms\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting SubForms"
If fso.FolderExists(sExportPath & "\Forms\SubForms\") Then
    Set delFold = fso.getfolder(sExportPath & "\Forms\SubForms\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting Modules"
If fso.FolderExists(sExportPath & "\Modules\") Then
    Set delFold = fso.getfolder(sExportPath & "\Modules\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting Macros"
If fso.FolderExists(sExportPath & "\Macros\") Then
    Set delFold = fso.getfolder(sExportPath & "\Macros\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting Reports"
If fso.FolderExists(sExportPath & "\Reports\") Then
    Set delFold = fso.getfolder(sExportPath & "\Reports\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting SubReports"
If fso.FolderExists(sExportPath & "\Reports\SubReports\") Then
    Set delFold = fso.getfolder(sExportPath & "\Reports\SubReports\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting Queries"
If fso.FolderExists(sExportPath & "\Queries\") Then
    Set delFold = fso.getfolder(sExportPath & "\Queries\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

addNote "  --Deleting SubQueries"
If fso.FolderExists(sExportPath & "\Queries\SubQueries\") Then
    Set delFold = fso.getfolder(sExportPath & "\Queries\SubQueries\")
    For Each delFile In delFold.Files
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

Set delFile = Nothing
Set delFold = Nothing

'---FORMS---
For Each myObj In oApplication.CurrentProject.AllForms
    If Not fso.FolderExists(sExportPath & "\Forms\") Then MkDir (sExportPath & "\Forms\")
    addNote "  " & myObj.FullName
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
    addNote "  " & myObj.FullName
    oApplication.SaveAsText acModule, myObj.FullName, sExportPath & "\Modules\" & myObj.FullName & ".bas"
Next

For Each myObj In oApplication.CurrentProject.AllMacros
    If Not fso.FolderExists(sExportPath & "\Macros\") Then MkDir (sExportPath & "\Macros\")
    addNote "  " & myObj.FullName
    oApplication.SaveAsText acMacro, myObj.FullName, sExportPath & "\Macros\" & myObj.FullName & ".mod"
Next

'---REPORTS---
For Each myObj In oApplication.CurrentProject.AllReports
    If Not fso.FolderExists(sExportPath & "\Reports\") Then MkDir (sExportPath & "\Reports\")
    addNote "  " & myObj.FullName
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
        addNote "  " & myObj.Name
        If Left(myObj.Name, 1) = "s" Then
            If Not fso.FolderExists(sExportPath & "\Queries\SubQueries\") Then MkDir (sExportPath & "\Queries\SubQueries\")
            oApplication.SaveAsText acQuery, myObj.Name, sExportPath & "\Queries\SubQueries\" & myObj.Name & ".qry"
        Else
            oApplication.SaveAsText acQuery, myObj.Name, sExportPath & "\Queries\" & myObj.Name & ".qry"
        End If
    End If
Next

Set myObj = Nothing
oApplication.CloseCurrentDatabase
oApplication.Quit
Set oApplication = Nothing
Set fso = Nothing

MsgBox "Files Decomposed from " & sADPFilename, vbInformation, "Nicely Done"

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
Form_frmTracking.Requery

moveTrackingToLastRecord

DoEvents

End Function