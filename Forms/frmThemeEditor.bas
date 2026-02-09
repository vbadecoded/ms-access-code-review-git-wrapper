Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function applyThemeChanges()

'All the theme information is in TEMPVARS so it resets when you close it and it will persist an entire database session. This could be a local session variables table as well
TempVars.Add "themePrimary", Me.primaryColor.Value
TempVars.Add "themeSecondary", Me.secondaryColor.Value
TempVars.Add "themeAccent", Me.accentColor.Value

If Me.darkMode Then
    TempVars.Add "themeMode", "Dark"
Else
    TempVars.Add "themeMode", "Light"
End If

TempVars.Add "themeColorLevels", Me.colorLevels.Value

'trying to prevent flashing...
DoCmd.Hourglass True
Me.Painting = False
DoCmd.Echo False

'This code applies the theme to ALL open forms

Dim f As Form, sForm As Control
Dim i As Integer

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

Call setTheme(Me)
Call setTheme(Me.sfrmThemeEditor.Form)

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor
Me.showAccent.BackColor = Me.accentColor

'make sure the form updates again
DoCmd.Hourglass False
Me.Painting = True
DoCmd.Echo True

End Function

Private Sub accentColor_Click()

If Me.Dirty Then Me.Dirty = False
Me.ActiveControl = colorPicker(Me.ActiveControl)

'Me.showPrimary.BackColor = Me.primaryColor
'Me.showSecondary.BackColor = Me.secondaryColor
Me.showAccent.BackColor = Me.accentColor

applyThemeChanges

End Sub

Private Sub colorLevels_AfterUpdate()

splitColorArray

End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor

End Sub

Private Sub Form_Load()

Call setTheme(Me)

splitColorArray
    
End Sub

Function applyLevels()

Select Case ""
    Case Nz(Me.L1), Nz(Me.L2), Nz(Me.L3), Nz(Me.L4)
        Exit Function
    Case Else
        Me.colorLevels = Me.L1 & "," & Me.L2 & "," & Me.L3 & "," & Me.L4
        applyThemeChanges
End Select

End Function

Public Function splitColorArray()

Dim splitIt() As String

splitIt = Split(Me.colorLevels, ",")

Me.L1 = splitIt(0)
Me.L2 = splitIt(1)
Me.L3 = splitIt(2)
Me.L4 = splitIt(3)

End Function

Private Sub L1_AfterUpdate()

applyLevels

End Sub

Private Sub L2_AfterUpdate()

applyLevels

End Sub

Private Sub L3_AfterUpdate()

applyLevels

End Sub

Private Sub L4_AfterUpdate()

applyLevels

End Sub

Private Sub newTheme_Click()

DoCmd.GoToRecord , , acNewRec

End Sub

Private Sub primaryColor_Click()

If Me.Dirty Then Me.Dirty = False
Me.ActiveControl = colorPicker(Me.ActiveControl)

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor

applyThemeChanges

End Sub

Private Sub secondaryColor_Click()

If Me.Dirty Then Me.Dirty = False
Me.ActiveControl = colorPicker(Me.ActiveControl)

Me.showPrimary.BackColor = Me.primaryColor
Me.showSecondary.BackColor = Me.secondaryColor

applyThemeChanges

End Sub

Private Sub testTheme_Click()

If Me.Dirty Then Me.Dirty = False
applyThemeChanges

End Sub
