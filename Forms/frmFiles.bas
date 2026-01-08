Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub runCommand_Click()

If Me.Dirty Then Me.Dirty = False

Dim db As Database
Set db = CurrentDb()

Dim rs As Recordset
Set rs = db.OpenRecordset("SELECT * FROM tblFiles WHERE selected = TRUE")

Dim fileList As String
fileList = ""

Do While Not rs.EOF
    fileList = fileList & " " & Replace(Replace(rs!location, Chr(9), ""), Chr(32), "")
    rs.MoveNext
Loop


Select Case Me.gitCmd
    Case "git diff"
        addNote "git diff "
        Call runGitCmd("git diff" & fileList)
    Case "git commit"
        addNote "git add" & fileList
        Call runGitCmd("git add" & fileList)
        DoEvents
        addNote "git commit -m """ & Form__MAIN.releaseNotes & """"
        Call runGitCmd("git commit -m " & Form__MAIN.releaseNotes)
    Case "recompose"
        addNote "Recomposing Files"
        Call recomposeAccdb(Form__MAIN.cmdRepo)
End Select

rs.Close
Set rs = Nothing
Set db = Nothing

End Sub

Private Sub selectAll_Click()

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tblFiles SET selected = TRUE"
DoCmd.SetWarnings True
Me.Requery

End Sub

Private Sub unSelectAll_Click()

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tblFiles SET selected = False"
DoCmd.SetWarnings True
Me.Requery

End Sub
