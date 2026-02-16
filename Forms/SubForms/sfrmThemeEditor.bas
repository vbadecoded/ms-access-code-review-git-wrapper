Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Paint()
On Error Resume Next

Me.primaryColor.BackColor = Me.primaryColor
Me.primaryColor.ForeColor = Me.primaryColor

If Me.secondaryColor = 0 Then
    Me.secondaryColor.BackColor = Me.primaryColor
    Me.secondaryColor.ForeColor = Me.primaryColor
Else
    Me.secondaryColor.BackColor = Me.secondaryColor
    Me.secondaryColor.ForeColor = Me.secondaryColor
End If

If Me.darkMode Then
    Me.dMode.BackColor = 0
    Me.dMode.ForeColor = vbWhite
    Me.themeName.ForeColor = vbWhite
Else
    Me.dMode.BackColor = vbWhite
    Me.dMode.ForeColor = 0
    Me.themeName.ForeColor = 0
End If

End Sub

Private Sub Form_Load()

Call setTheme(Me)
    
End Sub

Private Sub themeName_Click()

Dim f As Form, sForm As Control
Dim i As Integer

TempVars.Add "themePrimary", Me.primaryColor.value
TempVars.Add "themeSecondary", Me.secondaryColor.value
TempVars.Add "themeAccent", Me.accentColor.value

If Me.darkMode Then
    TempVars.Add "themeMode", "Dark"
Else
    TempVars.Add "themeMode", "Light"
End If

TempVars.Add "themeColorLevels", Me.colorLevels.value

Form_frmThemeEditor.filter = "recordId = " & Me.recordId
Form_frmThemeEditor.FilterOn = True

Call assignThemeToParameters(Me.recordId)

Form_frmThemeEditor.splitColorArray
DoEvents

Dim obj

For Each obj In Application.CurrentProject.AllForms
    If obj.IsLoaded = False Then GoTo nextOne
    Set f = forms(obj.Name)
    Call setTheme(f)
    For Each sForm In f.Controls
        If sForm.ControlType = acSubform Then
            On Error Resume Next
            Call setTheme(sForm.Form)
        End If
    Next sForm
nextOne:
Next obj

End Sub
