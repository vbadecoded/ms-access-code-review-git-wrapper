Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub location_Click()

Dim gitCmd As String

If Me.fileStatus = "staged" Then
    gitCmd = "git diff --cached "
Else
    gitCmd = "git diff "
End If

addNote gitCmd & Me.location

'add all modified files
Dim results As String
results = runGitCmd(gitCmd & Trim(Me.location), printNone:=True)

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

Private Sub stage_Click()

If Me.fileStatus <> "staged" Then
    Call runGitCmd("git add " & Me.location)
    Call Form__MAIN.gitStatus_Click
Else
    Call runGitCmd("git reset " & Me.location)
    Call Form__MAIN.gitStatus_Click
End If


End Sub
