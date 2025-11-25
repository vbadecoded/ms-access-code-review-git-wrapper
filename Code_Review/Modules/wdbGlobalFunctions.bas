Option Compare Database
Option Explicit

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal lpnShowCmd As Long) As Long

Public Sub openPath(Path)
CreateObject("Shell.Application").open CVar(Path)
End Sub

Public Function registerWdbUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String = "", Optional tag1 As Variant = "")
Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("tblWdbUpdateTracking")

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)

With rs1
    .AddNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = CStr(Nz(oldVal, ""))
        !newData = CStr(Nz(newVal, ""))
        !dataTag0 = tag0
        !dataTag1 = tag1
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing
End Function

Function emailContentGen(subject As String, Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String) As String
emailContentGen = subject & "," & Title & "," & subTitle & "," & primaryMessage & "," & detail1 & "," & detail2 & "," & detail3
End Function

Function userData(data) As String

Dim db As Database
Set db = OpenDatabase("\\data\mdbdata\WorkingDB\build\Repo\WorkingDB_Connection\WorkingDB_Connection.accdb")

Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & Environ("username") & "'")
userData = Nz(rsPermissions(data))
rsPermissions.Close

db.Close
End Function

Function getEmail(userName As String) As String

Dim db As Database
Set db = OpenDatabase("\\data\mdbdata\WorkingDB\_docs\Reporting\WorkingDB_ForExcel.accdb")

Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & userName & "'")
getEmail = rsPermissions!userEmail
rsPermissions.Close

db.Close

End Function

Function ap_DisableShift()

On Error GoTo errDisableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

db.Properties("AllowByPassKey") = False
Exit Function

errDisableShift:
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, False)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function

Public Function StrQuoteReplace(strValue)

StrQuoteReplace = Replace(Nz(strValue, ""), "'", "''")

End Function

Function ap_EnableShift()

On Error GoTo errEnableShift
Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()
db.Properties("AllowByPassKey") = True
Exit Function

errEnableShift:
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, True)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function