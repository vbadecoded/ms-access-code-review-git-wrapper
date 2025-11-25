Option Compare Database

Function disableShift()

Dim db, acc
Set acc = CreateObject("Access.Application")
'Set db = acc.DBEngine.OpenDatabase("\\data\mdbdata\WorkingDB\build\Commands\Misc_Commands\WorkingDB_SummaryEmail.accdb", False, False)
'Set db = acc.DBEngine.OpenDatabase("H:\dev\WorkingDB_SummaryEmail.accdb", False, False)
Set db = acc.DBEngine.OpenDatabase("C:\workingdb\WorkingDB_ghost.accdb", False, False)


db.Properties("AllowByPassKey") = True

db.Close
Set db = Nothing

End Function

Function disableShift_FE()

Dim db, acc
Set acc = CreateObject("Access.Application")

Dim repoLoc As String
repoLoc = currentRepoLocation & "Front_End\WorkingDB_FE.accdb"

Set db = acc.DBEngine.OpenDatabase(repoLoc, False, False)

db.Properties("AllowByPassKey") = True

db.Close
Set db = Nothing

End Function

Function getPassword()

Dim db As Database
Set db = OpenDatabase("")

Dim rs As Recordset
Set rs = db.OpenRecordset("SELECT * FROM MSysObjects WHERE Connect is not null")

Do While Not rs.EOF
    Debug.Print "Database: " & rs!Database & vbTab & " Connection: " & rs!Connect
    rs.MoveNext
Loop

rs.Close
Set rs = Nothing
db.Close
Set db = Nothing

End Function

Function currentRepoLocation() As String

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

currentRepoLocation = fso.GetParentFolderName(CurrentProject.Path) & "\"

Set fso = Nothing

End Function