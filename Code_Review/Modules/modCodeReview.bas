Option Compare Database
Option Explicit

Function runGitCmd(inputCmd As String, Optional dir As String = "current") As String

Dim wsShell As Object
Dim execObject As Object
Dim sOutput As String
Dim sWorkingDirectory As String

' Set the working directory to your Git repository
If dir = "current" Then
    sWorkingDirectory = currentRepoLocation
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

Dim strOutput
With CreateObject("Scripting.FileSystemObject")
    strOutput = .OpenTextFile(Environ("temp") & "\tempgitoutput.txt").ReadAll()
    .DeleteFile Environ("temp") & "\tempgitoutput.txt"
End With

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

Function decomposeAccdb(sADPFilename As String, sExportPath As String)

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim myType, myName, myPath, sStubADPFilename As String
myType = fso.GetExtensionName(sADPFilename)
myName = fso.GetBaseName(sADPFilename)
myPath = fso.GetParentFolderName(sADPFilename)

sStubADPFilename = sExportPath & "\" & myName & "_stub." & myType
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
If fso.FolderExists(sExportPath & "\Forms\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Forms\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Forms\SubForms\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Forms\SubForms\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Modules\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Modules\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Macros\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Macros\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Reports\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Reports\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Reports\SubReports\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Reports\SubReports\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Queries\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Queries\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

If fso.FolderExists(sExportPath & "\Queries\SubQueries\") Then
    Set delFold = fso.GetFolder(sExportPath & "\Queries\SubQueries\")
    For Each delFile In delFold.Files
        addNote "  " & delFile.Path
        fso.DeleteFile delFile.Path, True ' True for force deletion
    Next
End If

Set delFold = Nothing

'---FORMS---
For Each myObj In oApplication.CurrentProject.AllForms
    addNote "  " & myObj.FullName
    'move all new files
    If Left(myObj.FullName, 1) = "s" Then
        oApplication.SaveAsText acForm, myObj.FullName, sExportPath & "\Forms\SubForms\" & myObj.FullName & ".form"
    Else
        oApplication.SaveAsText acForm, myObj.FullName, sExportPath & "\Forms\" & myObj.FullName & ".form"
    End If
Next

'---MODULES---
For Each myObj In oApplication.CurrentProject.AllModules
    addNote "  " & myObj.FullName
    oApplication.SaveAsText acModule, myObj.FullName, sExportPath & "\Modules\" & myObj.FullName & ".bas"
Next

For Each myObj In oApplication.CurrentProject.AllMacros
    addNote "  " & myObj.FullName
    oApplication.SaveAsText acMacro, myObj.FullName, sExportPath & "\Macros\" & myObj.FullName & ".mod"
Next

'---REPORTS---
For Each myObj In oApplication.CurrentProject.AllReports
    addNote "  " & myObj.FullName
    If Left(myObj.FullName, 1) = "s" Then
        oApplication.SaveAsText acReport, myObj.FullName, sExportPath & "\Reports\SubReports\" & myObj.FullName & ".rpt"
    Else
        oApplication.SaveAsText acReport, myObj.FullName, sExportPath & "\Reports\" & myObj.FullName & ".rpt"
    End If
Next

'---QUERIES---
For Each myObj In oApplication.CurrentDb.QueryDefs
    If Not Left(myObj.Name, 3) = "~sq" Then 'exclude queries defined by the forms. Already included in the form itself
        addNote "  " & myObj.Name
        If Left(myObj.Name, 1) = "s" Then
            oApplication.SaveAsText acQuery, myObj.Name, sExportPath & "\Queries\SubQueries\" & myObj.Name & ".qry"
        Else
            oApplication.SaveAsText acQuery, myObj.Name, sExportPath & "\Queries\" & myObj.Name & ".qry"
        End If
    End If
Next

oApplication.CloseCurrentDatabase
oApplication.Quit

fso.DeleteFile sStubADPFilename

MsgBox "Files Decomposed from " & sADPFilename, vbInformation, "Nicely Done"

End Function

Function addNote(noteTxt As String)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO tblReleaseTracking(task) VALUES('" & StrQuoteReplace(noteTxt) & " ')"
DoCmd.SetWarnings True

On Error Resume Next
Form_frmTracking.Requery

DoEvents

End Function