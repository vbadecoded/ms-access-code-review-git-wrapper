Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub location_Click()
addNote "git diff " & Me.location

'add all modified files
Dim results As String
results = runGitCmd("git diff " & Trim(Me.location))

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * from tblDiff"
DoCmd.SetWarnings True

Dim arr() As String
arr = Split(results, vbLf)
DoCmd.SetWarnings False
Dim ITEM
For Each ITEM In arr
        DoCmd.RunSQL "INSERT INTO tblDiff(diffLine) VALUES('" & StrQuoteReplace(ITEM) & "')"
Next ITEM
DoCmd.SetWarnings True

Form__MAIN.lblGitDiff.Caption = "Git Diff " & Me.location
Form__MAIN.sfrmDiff.Requery
End Sub

Private Sub selectAll_Click()

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tblFiles SET selected = TRUE"
DoCmd.SetWarnings True
Me.Requery

End Sub

Private Sub stage_Click()

Call runGitCmd("git add " & Me.location)
Call Form__MAIN.gitStatus_Click

End Sub

Private Sub unSelectAll_Click()

DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE tblFiles SET selected = False"
DoCmd.SetWarnings True
Me.Requery

End Sub
