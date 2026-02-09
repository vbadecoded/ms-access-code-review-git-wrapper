Option Compare Database
Option Explicit

Function setSplashLoading(label As String)

If IsNull(TempVars!loadAmount) Then Exit Function
TempVars.Add "loadAmount", TempVars!loadAmount + 1
Form_frmSplash.lnLoading.Width = (TempVars!loadAmount / 5) * TempVars!loadWd
Form_frmSplash.lblLoading.Caption = label
Form_frmSplash.Repaint

End Function

Function assignThemeToParameters(themeId As Long)

Dim db As Database
Set db = CurrentDb

db.Execute "UPDATE tblParameters SET themeId = " & themeId

Set db = Nothing

End Function

Function disableShift()

Dim db, acc
Set acc = CreateObject("Access.Application")
'Set db = acc.DBEngine.OpenDatabase("\\data\mdbdata\WorkingDB\build\Commands\Misc_Commands\WorkingDB_SummaryEmail.accdb", False, False)
'Set db = acc.DBEngine.OpenDatabase("H:\dev\WorkingDB_SummaryEmail.accdb", False, False)
Set db = acc.DBEngine.OpenDatabase("C:\workingdb\WorkingDB_ghost.accdb", False, False)


db.Properties("AllowByPassKey") = True

db.CLOSE
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

rs.CLOSE
Set rs = Nothing
db.CLOSE
Set db = Nothing

End Function