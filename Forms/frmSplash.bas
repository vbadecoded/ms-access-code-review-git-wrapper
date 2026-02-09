Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

TempVars.Add "loadAmount", 0
TempVars.Add "loadWd", 8160

Me.lblFrozen.Visible = False
Call setSplashLoading("Loading.")

DoEvents
Form_frmSplash.SetFocus
DoEvents

Call setSplashLoading("Loading..")

If CommandBars("Ribbon").Height > 100 Then CommandBars.ExecuteMso "MinimizeRibbon"

Call setSplashLoading("Loading...")

'set up theme
Dim themeId As Long
themeId = Nz(DLookup("themeId", "tblParameters"), 0)

Dim db As Database
Set db = CurrentDb()

Dim rsTheme As Recordset

If themeId <> 0 Then
    Set rsTheme = db.OpenRecordset("SELECT * FROM tblTheme WHERE recordId = " & themeId)
    
    If rsTheme!darkMode.Value Then
        TempVars.Add "themeMode", "Dark"
    Else
        TempVars.Add "themeMode", "Light"
    End If
    
    TempVars.Add "themePrimary", CStr(rsTheme!primaryColor.Value)
    TempVars.Add "themeSecondary", CStr(rsTheme!secondaryColor.Value)
    TempVars.Add "themeAccent", CStr(rsTheme!accentColor.Value)
    TempVars.Add "themeColorLevels", CStr(rsTheme!colorLevels.Value)
    
    rsTheme.CLOSE
    Set rsTheme = Nothing
End If

Set db = Nothing

Call setSplashLoading("Loading....")

DoCmd.OpenForm "_MAIN"
Form__MAIN.Visible = False

Call setSplashLoading("Loading.....")

DoCmd.CLOSE acForm, "frmSplash"
DoEvents
Form__MAIN.Visible = True
DoCmd.Maximize
    
End Sub
