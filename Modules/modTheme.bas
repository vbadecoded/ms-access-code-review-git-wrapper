Option Compare Database
Option Explicit

'---this is an API for the color picker---
'I use this on the frmThemeEditor to select colors in window
Declare PtrSafe Sub ChooseColor Lib "msaccess.exe" Alias "#53" (ByVal hwnd As LongPtr, rgb As Long)


Public Function setTheme(setForm As Form)
'this is not an error prone routine... but IF there are errors - this is not one I typically track in my production FEs.
'Feel free to add error trapping to this though. Could we worthwhile.
On Error Resume Next

Dim colorLevArr() As String

Dim scalarBack As Double, scalarFront As Double, darkMode As Boolean
Dim backBase As Long, foreBase As Long, backAccent As Long, colorLevels(4), backSecondary As Long, btnXback As Long, btnXbackShade As Long

Dim ctl As Control, eachBtn As CommandButton
Dim classColor As String, fadeBack, fadeFore
Dim Level
Dim backCol As Long, levFore As Double
Dim disFore As Double
Dim foreLevInt As Long, maxLev As Long

'IF NO THEME SET, APPLY DEFAULT THEME (for Dev mode mostly)
If Nz(TempVars!themePrimary, "") = "" Then
    TempVars.Add "themePrimary", 3355443
    TempVars.Add "themeSecondary", 0
    TempVars.Add "themeAccent", 5787704
    TempVars.Add "themeMode", "Dark"
    TempVars.Add "themeColorLevels", "1.3,1.6,1.9,2.2"
End If

darkMode = TempVars!themeMode = "Dark"

'set some manual values based on dark/light theme.
'the scalar values are somewhat arbitrary.
If darkMode Then
    foreBase = 16777215
    btnXback = 4342397
    scalarBack = 1.3
    scalarFront = 0.9
Else
    foreBase = 657930
    btnXback = 8947896
    scalarBack = 1.1
    scalarFront = 0.3
End If

'these are the raw base colors
backBase = CLng(TempVars!themePrimary)
backSecondary = CLng(TempVars!themeSecondary)
backAccent = CLng(TempVars!themeAccent)

'to achieve the 5 'Levels' of controls, this array is the primary method.
colorLevArr = Split(TempVars!themeColorLevels, ",")

If backSecondary <> 0 Then 'if the theme contains a primary AND a secondary color
    colorLevels(0) = backBase
    colorLevels(1) = shadeColor(backSecondary, CDbl(colorLevArr(0)))
    colorLevels(2) = shadeColor(backBase, CDbl(colorLevArr(1)))
    colorLevels(3) = shadeColor(backSecondary, CDbl(colorLevArr(2)))
    colorLevels(4) = shadeColor(backBase, CDbl(colorLevArr(3)))
Else 'if the theme only contains a primary color
    colorLevels(0) = backBase
    colorLevels(1) = shadeColor(backBase, CDbl(colorLevArr(0)))
    colorLevels(2) = shadeColor(backBase, CDbl(colorLevArr(1)))
    colorLevels(3) = shadeColor(backBase, CDbl(colorLevArr(2)))
    colorLevels(4) = shadeColor(backBase, CDbl(colorLevArr(3)))
End If

'set the form parts themes
setForm.FormHeader.BackColor = colorLevels(findColorLevel(setForm.FormHeader.Tag))
setForm.Detail.BackColor = colorLevels(findColorLevel(setForm.Detail.Tag))
If Len(setForm.Detail.Tag) = 4 Then
    setForm.Detail.AlternateBackColor = colorLevels(findColorLevel(setForm.Detail.Tag) + 1)
Else
    setForm.Detail.AlternateBackColor = setForm.Detail.BackColor
End If

setForm.FormFooter.BackColor = colorLevels(findColorLevel(setForm.FormFooter.Tag))
'NOTE - this does assume form parts don't use tags for other purposes


'---PRIMARY THEME SETTING AREA---
'a giant For Each with Select Cases. Not rocket science.


For Each ctl In setForm.Controls 'simply loop through all controls on the form
    If Not ctl.Tag Like "*.L#*" Then GoTo nextControl 'is there a tag with a theme attribute on it? if not - skip this control
    
    '---
    '---FOR ALL CONTROLS---
    Level = findColorLevel(ctl.Tag)
    backCol = colorLevels(Level)
    foreLevInt = Level
    If foreLevInt > 3 Then foreLevInt = 3
    
    If darkMode Then
        levFore = (1 / colorLevArr(foreLevInt)) + 0.2
        disFore = 1.4 - levFore
    Else
        levFore = (colorLevArr(foreLevInt))
        disFore = 15 - levFore
    End If
    
    maxLev = Level + 1
    If maxLev > 4 Then maxLev = 4
    If ctl.Tag Like "*ContrastBorder*" Then
        ctl.BorderColor = colorLevels(maxLev)
    Else
        ctl.BorderColor = backCol
    End If
    
    '--now, find the control type and apply the applicable
    Select Case ctl.ControlType
        '---
        '---COMMAND BUTTON
        Case acCommandButton, acToggleButton
            ctl.BackColor = backCol
            
            '---this is for swapping out button icons for light / dark theme icons - turned off by default---
            '            If (ctl.Picture = "") Then GoTo skipAhead0
            '            If darkMode Then
            '                If InStr(ctl.Picture, "\Core_theme_light\") Then ctl.Picture = Replace(ctl.Picture, "\Core_theme_light\", "\Core\")
            '            Else
            '                If InStr(ctl.Picture, "\Core\") Then ctl.Picture = Replace(ctl.Picture, "\Core\", "\Core_theme_light\")
            '            End If
            '---
            
            
            '---test for individual attributes---
            
            If ctl.Tag Like "*dis*" Then
                fadeFore = shadeColor(foreBase, disFore)
                ctl.ForeColor = fadeFore
                ctl.HoverForeColor = fadeFore
                ctl.PressedForeColor = fadeFore
            Else
                fadeFore = shadeColor(foreBase, levFore - 0.2)
                ctl.ForeColor = foreBase
                ctl.HoverForeColor = foreBase
                ctl.PressedForeColor = foreBase
            End If
            
            If ctl.Tag Like "*btnX*" Then
                fadeBack = shadeColor(btnXback, scalarBack)
                btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                ctl.BackColor = btnXbackShade
                ctl.BorderColor = btnXback
            Else
                fadeBack = shadeColor(backCol, scalarBack)
            End If
            
            ctl.HoverColor = fadeBack
            ctl.PressedColor = fadeBack
            
            If ctl.Tag Like "*cardBtn*" Then
                ctl.HoverColor = backCol
                ctl.PressedColor = backCol
            End If
            
            If ctl.Tag Like "*accentBtn*" Then
                fadeBack = shadeColor(backAccent, scalarBack)
                ctl.BackColor = backAccent
                ctl.Gradient = 17
            End If
        '---
        '---LABEL
        Case acLabel
            ctl.ForeColor = shadeColor(foreBase, levFore)
            If ctl.Tag Like "*lbl_wBack.L#*" Then ctl.BackColor = backCol
        '---
        '---TEXT BOX
        Case acTextBox, acComboBox
            ctl.BackColor = backCol
            If ctl.Tag Like "*txtTransFore*" Then
                ctl.ForeColor = backCol
            ElseIf ctl.Tag Like "*txtErr*" Then
                ctl.BorderColor = btnXback
                ctl.BorderStyle = 1
                ctl.ForeColor = foreBase
            Else
                ctl.ForeColor = foreBase
            End If
            
            If ctl.FormatConditions.Count = 1 Then 'special case for null value conditional formatting. Typically this is used for placeholder values
                If ctl.FormatConditions.item(0).Expression1 Like "*IsNull*" Then
                    ctl.FormatConditions.item(0).BackColor = backCol
                    ctl.FormatConditions.item(0).ForeColor = foreBase
                End If
            End If
        '---
        '---BOX / SUBFORM
        Case acRectangle, acSubform
            If Not ctl.Name Like "sfrm*" Then ctl.BackColor = backCol
        '---
        '---TAB CONTROL
        Case acTabCtl
            ctl.PressedColor = backCol
            fadeBack = shadeColor(CLng(colorLevels(Level - 1)), scalarBack)
            ctl.HoverColor = fadeBack
            ctl.HoverForeColor = foreBase
            ctl.PressedForeColor = foreBase
            If Level = 0 Then
                ctl.BackColor = colorLevels(Level + 0)
                fadeFore = shadeColor(foreBase, levFore - 0.6)
                ctl.ForeColor = fadeFore
            Else
                ctl.BackColor = colorLevels(Level - 1)
                fadeFore = shadeColor(foreBase, levFore)
                ctl.ForeColor = fadeFore
            End If
        '---
        '---PICTURE
        Case acImage
            ctl.BackColor = backCol
    End Select
    
nextControl:
Next

End Function

Function findColorLevel(tagText As String) As Long

findColorLevel = 0
If tagText = "" Then Exit Function

findColorLevel = Mid(tagText, InStr(tagText, ".L") + 2, 1)

End Function

Function shadeColor(inputColor As Long, scalar As Double) As Long

Dim tempHex, ioR, ioG, ioB

tempHex = Hex(inputColor)

If tempHex = "0" Then tempHex = "111111"

If Len(tempHex) = 1 Then tempHex = "0" & tempHex
If Len(tempHex) = 2 Then tempHex = "0" & tempHex
If Len(tempHex) = 3 Then tempHex = "0" & tempHex
If Len(tempHex) = 4 Then tempHex = "0" & tempHex
If Len(tempHex) = 5 Then tempHex = "0" & tempHex

ioR = Val("&H" & Mid(tempHex, 5, 2)) * scalar
ioG = Val("&H" & Mid(tempHex, 3, 2)) * scalar
ioB = Val("&H" & Mid(tempHex, 1, 2)) * scalar

'Debug.Print ioR & " "; ioG & " " & ioB

If ioR > 255 Then ioR = 255
If ioG > 255 Then ioG = 255
If ioB > 255 Then ioB = 255

If ioR < 0 Then ioR = 0
If ioG < 0 Then ioG = 0
If ioB < 0 Then ioB = 0

shadeColor = rgb(ioR, ioG, ioB)

End Function

Public Function colorPicker(Optional lngColor As Long) As Long

    'Static lngColor As Long
    ChooseColor Application.hWndAccessApp, lngColor
    colorPicker = lngColor

End Function