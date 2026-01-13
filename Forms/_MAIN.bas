Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const acCmdCloseAll = &H286
Const acCmdCompileAndSaveAllModules = &H7E
Const msoControlButton = 1
Const acForm = 2
Const acModule = 5
Const acMacro = 4
Const acReport = 3
Const acQuery = 1

Dim fso As Object

Private Sub clearLog_Click()
formStatus (True)

Dim db As Database
Set db = CurrentDb()

db.Execute "DELETE * from tblReleaseTracking"

Me.frmTracking.Requery

Set db = Nothing

formStatus (False)
End Sub

Private Sub cmdRepo_AfterUpdate()
formStatus (True)

Call getRepoInfo(Me.cmdRepo)

formStatus (False)
End Sub

Private Sub commitAll_Click()
formStatus (True)

'add all modified files
Dim results As String

addNote "git commit -a -m """ & Form__MAIN.releaseNotes & """"
Call runGitCmd("git commit -a -m """ & Form__MAIN.releaseNotes & """")

formStatus (False)
End Sub

Private Sub decompose_Click()
formStatus (True)

Dim filePath As String
Dim fileName As String

filePath = Form__MAIN.cmdRepo
fileName = Me.cmdRepo.Column(2)

Call decomposeAccdb(filePath & fileName, filePath)

formStatus (False)
End Sub

Private Sub disableShift_Click()
formStatus (True)

If shiftKeyBypass(getDB, False) Then addNote Me.cmdRepo.Column(2) & " Shift Key Disabled"

formStatus (False)
End Sub

Private Sub enableShift_Click()
formStatus (True)

If shiftKeyBypass(getDB, True) Then addNote Me.cmdRepo.Column(2) & " Shift Key Enabled"

formStatus (False)
End Sub

Private Sub Form_Load()
formStatus (True)

'initial data based on environment variables
Me.responsiblePerson = Environ("username")
Me.userEmail = getEmail(Environ("username"))

'----------------------------
'---REPOSITORY SEARCHING
'----------------------------

Set fso = CreateObject("Scripting.FileSystemObject")

Dim db As Database
Dim rsRepos As Recordset, rsFindRepo As Recordset

Set db = CurrentDb()
Set rsRepos = db.OpenRecordset("tblRepoLocation")

'first delete all records in rsRepo
Do While Not rsRepos.EOF
    rsRepos.Delete
    rsRepos.MoveNext
Loop
rsRepos.MoveFirst

'use FSO to scan folders near this repository - these are treated as repositories to work on IF an .accdb or .mdb file is found
Dim f, sf, sfo
Set f = fso.getfolder(fso.getparentfoldername(CurrentProject.Path))
Set sf = f.subfolders

Dim fsDB, fsDBName As String, fsProdLocName As String
For Each sfo In sf
    'look for the record first - skip if found
    If rsRepos.RecordCount > 0 Then
        Set rsFindRepo = db.OpenRecordset("SELECT * FROM tblRepoLocation WHERE repoLocation = '" & sfo.Path & "\" & "'")
        If rsFindRepo.RecordCount > 0 Then GoTo skipRepo
    End If
    
    'now scan for the .accdb/.mdb and skip if not found
    fsDBName = ""
    fsProdLocName = ""
    For Each fsDB In sfo.Files
        Select Case fso.GetExtensionName(fsDB.Path)
            Case "accdb", "mdb"
                'database found!
                fsDBName = fsDB.Name
            Case "txt"
                If fsDB.Name = ".productionLocation.txt" Then
                    'Get first line of text document
                    Open fsDB.Path For Input As #1
                    Line Input #1, fsProdLocName
                    Close #1
                End If
        End Select
    Next fsDB
    
    If fsDBName = "" Then GoTo skipRepo 'no database found
    
    rsRepos.AddNew
    rsRepos!repoLocation = sfo.Path & "\"
    rsRepos!dbName = fsDBName
    rsRepos!productionLocation = fsProdLocName
    rsRepos.Update
    
skipRepo:
Next

'----------------------------
'---REPOSITORY SELECTION
'----------------------------

'---AUTO SELECT REPO IF ONLY ONE---
If rsRepos.RecordCount = 1 Then
    Me.cmdRepo = rsRepos!repoLocation & "\"
    Call getRepoInfo(rsRepos!repoLocation & "\")
Else
'---IF MORE THAN ONE REPO FOUND, CHECK tblLastUsed---
    'if more than one is found, check if the previously used repo is available in the LastUsed table
    Dim rsLU As Recordset
    Set rsLU = db.OpenRecordset("SELECT * from tblLastUsed WHERE recordId = 1")
    If Nz(rsLU!repoLocation, "") = "" Then GoTo LUnotFound 'blank field
    Set rsFindRepo = db.OpenRecordset("SELECT * FROM tblRepoLocation WHERE repoLocation = '" & rsLU!repoLocation & "'")
    If rsFindRepo.RecordCount = 1 Then 'last used repo found!!
        Me.cmdRepo = rsLU!repoLocation
        Call getRepoInfo(rsLU!repoLocation)
    End If
    
LUnotFound:
End If

'---Cleanup---
On Error Resume Next
rsLU.Close: Set rsLU = Nothing
rsFindRepo.Close: Set rsFindRepo = Nothing
rsRepos.Close: Set rsRepos = Nothing
Set db = Nothing

Set fso = Nothing


DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tblReleaseTracking"
DoCmd.RunSQL "INSERT INTO tblReleaseTracking(task) values('Form Initialized')"
DoCmd.SetWarnings True
Me.frmTracking.Requery

formStatus (False)
End Sub

Private Sub gitbranch_AfterUpdate()
formStatus (True)
addNote "git checkout " & Me.gitbranch

Call runGitCmd("git checkout " & Me.gitbranch)

formStatus (False)
End Sub

Private Sub gitCommit_Click()
formStatus (True)

'add all modified files
Dim results As String
results = runGitCmd("git status")

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * from tblFiles"
DoCmd.SetWarnings True

Dim arr() As String
arr = Split(results, vbLf)
DoCmd.SetWarnings False
Dim item
For Each item In arr
    If InStr(item, "modified:") Then DoCmd.RunSQL "INSERT INTO tblFiles(location) VALUES('" & Trim(Replace(item, "modified:", "")) & " ')"
Next item
DoCmd.SetWarnings True
DoCmd.OpenForm "frmFiles"
Form_frmFiles.gitCmd = "git commit"

formStatus (False)
End Sub

Private Sub gitDiff_Click()
formStatus (True)
'add all modified files
Dim results As String
results = runGitCmd("git status")

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * from tblFiles"
DoCmd.SetWarnings True

Dim arr() As String
arr = Split(results, vbLf)
DoCmd.SetWarnings False
Dim item
For Each item In arr
    If InStr(item, "modified:") Then DoCmd.RunSQL "INSERT INTO tblFiles(location) VALUES('" & Trim(Replace(item, "modified:", "")) & " ')"
Next item
DoCmd.SetWarnings True
DoCmd.OpenForm "frmFiles"
Form_frmFiles.gitCmd = "git diff"

formStatus (False)
End Sub

Private Sub gitMerge_Click()
formStatus (True)
addNote "git merge " & Me.gitBranchSelect

Call runGitCmd("git merge " & Me.gitBranchSelect)

formStatus (False)
End Sub

Private Sub gitPull_Click()
formStatus (True)
addNote "git pull origin " & Me.gitbranch

Call runGitCmd("git pull origin " & Me.gitbranch)

formStatus (False)
End Sub

Private Sub gitPush_Click()
formStatus (True)
addNote "git push origin " & Me.gitbranch

Call runGitCmd("git push origin " & Me.gitbranch)

formStatus (False)
End Sub

Private Sub gitStatus_Click()
formStatus (True)
addNote "git status"

Call runGitCmd("git status")

formStatus (False)
End Sub

Private Sub increaseRev_Click()
formStatus (True)
Dim x, Y, major, minor, min, newMajor, newMinor, newMin

x = Me.releaseNum
Y = Replace(x, "REV", "")
major = Split(Y, ".")(0)
minor = Split(Y, ".")(1)
min = Split(Y, ".")(2)

If (min <> 99) Then
    newMajor = major
    newMinor = minor
    newMin = min + 1
    If newMin < 10 Then newMin = "0" & newMin
    GoTo done
End If
newMin = "00"

If (minor <> 9) Then
    newMajor = major
    newMinor = minor + 1
    GoTo done
End If
newMinor = 0
newMajor = major + 1

done:
Dim newRel As String
newRel = "REV" & newMajor & "." & newMinor & "." & newMin
Me.releaseNum = newRel
addNote "Rev Increased to " & newRel
formStatus (False)
End Sub

Function formStatus(inWork As Boolean)

If inWork Then
    Me.Detail.BackColor = RGB(50, 0, 0)
Else
    Me.Detail.BackColor = RGB(38, 38, 38)
End If

Me.codeRunning.Visible = inWork

End Function

Private Sub openAccdb_Click()
formStatus (True)

Call openPath(getDB)

addNote Me.cmdRepo.Column(2) & " Opened"

formStatus (False)
End Sub

Private Sub openGitGUI_Click()
formStatus (True)
addNote "git gui"

Call runGitCmd("git gui")

formStatus (False)
End Sub

Private Sub publishChanges_Click()
formStatus (True)

addNote "git pull origin master : " & Me.cmdRepo.Column(1)
Call runGitCmd("git pull origin master", Me.cmdRepo.Column(1))

formStatus (False)
End Sub

Private Sub publishFE_Click()
formStatus (True)

Dim errorMsg As New Collection
If Nz(Me.releaseNotes, "") = "" Then errorMsg.Add "Empty release notes"

If errorMsg.Count > 0 Then
    Dim msgContents As String, item
    msgContents = ""
    For Each item In errorMsg
        msgContents = msgContents & vbNewLine & item
    Next item
    MsgBox msgContents, vbInformation, "Please fix these issues: "
    GoTo exitThis
End If

If MsgBox("Are you sure? ", vbYesNo, "Just Checking") = vbNo Then GoTo exitThis

addNote "--- STARTING " & Me.cmdRepo.Column(2) & " CLEANING PROCEDURE ---"

'---Setup Variables
addNote "establishing variables..."

Set fso = CreateObject("Scripting.FileSystemObject")

TempVars.Add "releaseNum", Me.releaseNum.Value
TempVars.Add "releaseNotes", Replace(Me.releaseNotes.Value, "'", "''")
TempVars.Add "responsiblePerson", Me.responsiblePerson.Value
TempVars.Add "userEmail", Me.userEmail.Value
TempVars.Add "databaseName", Me.cmdRepo.Column(2)

Dim repoLoc As String
repoLoc = Form__MAIN.cmdRepo
TempVars.Add "devFile", getDB

Dim devBackup, devTemp, feFile

'---Open DEV to finalize---
'-if Front End
If Me.cmdRepo = "workingdb" Then
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

addNote Me.cmdRepo.Column(2) & " CLEANED"

exitThis:
formStatus (False)
End Sub

Private Sub publishNOTES_Click()
formStatus (True)

DoCmd.SetWarnings False
DoCmd.RunSQL "INSERT INTO tblReleaseNotes" & _
    "(DatabaseVersion,Notes,ReleaseDate,ReleasedBy,DatabaseName)" & _
    " VALUES" & _
    "('" & Me.releaseNum & "','" & Me.releaseNotes & "','" & Date & "','" & Me.responsiblePerson & "','" & Me.cmdRepo.Column(2) & "');"
DoCmd.SetWarnings True

Dim body, strValues
addNote "Generate notification email"
body = emailContentGen("New Version Published", Me.cmdRepo.Column(2) & " " & Me.releaseNum & " Published", "Notes: " & Replace(Me.releaseNotes, ",", ";"), "Responsible: " & responsiblePerson, "Releaser: " & Environ("username"), "", "")

If Environ("username") <> "brownj" Then
    strValues = "'brownj','brownj@us.nifco.com','" & Environ("username") & "','" & getEmail(Environ("username")) & "','" & Now() & "',1,1,'New Version Published','" & body & "','" & Now() & "'"
    DoCmd.RunSQL "INSERT INTO tblNotificationsSP(recipientUser,recipientEmail,senderUser,senderEmail,sentDate,notificationType,notificationPriority,notificationDescription,emailContent,readDate) VALUES(" & strValues & ");"
    addNote "Notification sent to brownj"
End If

If Environ("username") <> "georgemi" Then
    strValues = "'georgemi','georgemi@us.nifco.com','" & Environ("username") & "','" & getEmail(Environ("username")) & "','" & Now() & "',1,1,'New Version Published','" & body & "','" & Now() & "'"
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO tblNotificationsSP(recipientUser,recipientEmail,senderUser,senderEmail,sentDate,notificationType,notificationPriority,notificationDescription,emailContent,readDate) VALUES(" & strValues & ");"
    DoCmd.SetWarnings True
    addNote "Notification sent to georgemi"
End If

addNote "Version " & Me.releaseNum & " Notes Published Successfully"

formStatus (False)
End Sub

Private Sub recomposeSendFile_Click()
formStatus (True)
Set fso = CreateObject("Scripting.FileSystemObject")

'---ADD ALL CHANGES FILES TO LIST---
Dim results As String
results = runGitCmd("git status")

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * from tblFiles"
DoCmd.SetWarnings True

Dim arr() As String
arr = Split(results, vbLf)
DoCmd.SetWarnings False
Dim item
For Each item In arr
    If InStr(item, "modified:") Then DoCmd.RunSQL "INSERT INTO tblFiles(location) VALUES('" & Trim(Replace(item, "modified:", "")) & " ')"
Next item
DoCmd.SetWarnings True

'open form files
DoCmd.OpenForm "frmFiles"
Form_frmFiles.gitCmd = "Recompose"

formStatus (False)
End Sub

Private Sub releaseHelp_Click()
formStatus (True)
DoCmd.OpenForm "frmHelp"
addNote "Opened Help Form"
formStatus (False)
End Sub

Private Sub responsiblePerson_AfterUpdate()
formStatus (True)
If Me.Dirty Then Me.Dirty = False
Me.userEmail = getEmail(Me.responsiblePerson)
addNote "Populated User Email"
formStatus (False)
End Sub

Private Sub runGit_Click()
formStatus (True)
addNote Me.gitCmd

Call runGitCmd(Me.gitCmd)

formStatus (False)
End Sub
