Option Compare Database
Option Explicit

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal lpnShowCmd As Long) As Long

Public Sub openPath(Path)
CreateObject("Shell.Application").open CVar(Path)
End Sub

Function emailContentGen(subject As String, Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String) As String
emailContentGen = subject & "," & Title & "," & subTitle & "," & primaryMessage & "," & detail1 & "," & detail2 & "," & detail3
End Function

Function getEmail(userName As String) As String
On Error Resume Next

Dim db As Database
Set db = CurrentDb()

Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblDeveloperInfo WHERE user = '" & userName & "'")
getEmail = rsPermissions!Email
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