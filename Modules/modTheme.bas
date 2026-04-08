option compare database
option explicit

'---this is an api for the color picker---
'i use this on the frmthemeeditor to select colors in window
declare ptrsafe sub choosecolor lib "msaccess.exe" alias "#53" (byval hwnd as longptr, rgb as long)


public function settheme(setform as form)
'this is not an error prone routine... but if there are errors - this is not one i typically track in my production fes.
'feel free to add error trapping to this though. could we worthwhile.

dim colorlevarr() as string

dim scalarback as double, scalarfront as double, darkmode as boolean
dim backbase as long, forebase as long, backaccent as long, colorlevels(4), backsecondary as long, btnxback as long, btnxbackshade as long

dim ctl as control, eachbtn as commandbutton
dim classcolor as string, fadeback, fadefore
dim level
dim backcol as long, levfore as double
dim disfore as double
dim forelevint as long, maxlev as long

'if no theme set, apply default theme (for dev mode mostly)
if nz(tempvars!themeprimary, "") = "" then
    tempvars.add "themePrimary", 3355443
    tempvars.add "themeSecondary", 0
    tempvars.add "themeAccent", 5787704
    tempvars.add "themeMode", "Dark"
    tempvars.add "themeColorLevels", "1.3,1.6,1.9,2.2"
end if

darkmode = tempvars!thememode = "Dark"

'set some manual values based on dark/light theme.
'the scalar values are somewhat arbitrary.
if darkmode then
    forebase = 16777215
    btnxback = 4342397
    scalarback = 1.3
    scalarfront = 0.9
else
    forebase = 657930
    btnxback = 8947896
    scalarback = 1.1
    scalarfront = 0.3
end if

'these are the raw base colors
backbase = clng(tempvars!themeprimary)
backsecondary = clng(tempvars!themesecondary)
backaccent = clng(tempvars!themeaccent)

'to achieve the 5 'levels' of controls, this array is the primary method.
colorlevarr = split(tempvars!themecolorlevels, ",")

if backsecondary <> 0 then 'if the theme contains a primary and a secondary color
    colorlevels(0) = backbase
    colorlevels(1) = shadecolor(backsecondary, cdbl(colorlevarr(0)))
    colorlevels(2) = shadecolor(backbase, cdbl(colorlevarr(1)))
    colorlevels(3) = shadecolor(backsecondary, cdbl(colorlevarr(2)))
    colorlevels(4) = shadecolor(backbase, cdbl(colorlevarr(3)))
else 'if the theme only contains a primary color
    colorlevels(0) = backbase
    colorlevels(1) = shadecolor(backbase, cdbl(colorlevarr(0)))
    colorlevels(2) = shadecolor(backbase, cdbl(colorlevarr(1)))
    colorlevels(3) = shadecolor(backbase, cdbl(colorlevarr(2)))
    colorlevels(4) = shadecolor(backbase, cdbl(colorlevarr(3)))
end if

'set the form parts themes
on error resume next
setform.formheader.backcolor = colorlevels(findcolorlevel(setform.formheader.tag))

setform.detail.backcolor = colorlevels(findcolorlevel(setform.detail.tag))
if len(setform.detail.tag) = 4 then
    setform.detail.alternatebackcolor = colorlevels(findcolorlevel(setform.detail.tag) + 1)
else
    setform.detail.alternatebackcolor = setform.detail.backcolor
end if

setform.formfooter.backcolor = colorlevels(findcolorlevel(setform.formfooter.tag))
if err.number = 2462 then err.clear
'note - this does assume form parts don't use tags for other purposes

on error goto 0

'---primary theme setting area---
'a giant for each with select cases. not rocket science.

for each ctl in setform.controls 'simply loop through all controls on the form
    if not ctl.tag like "*.L#*" then goto nextcontrol 'is there a tag with a theme attribute on it? if not - skip this control
    
    '---
    '---for all controls---
    level = findcolorlevel(ctl.tag)
    backcol = colorlevels(level)
    forelevint = level
    if forelevint > 3 then forelevint = 3
    
    if darkmode then
        levfore = (1 / colorlevarr(forelevint)) + 0.2
        disfore = 1.4 - levfore
    else
        levfore = (colorlevarr(forelevint))
        disfore = 15 - levfore
    end if
    
    maxlev = level + 1
    if maxlev > 4 then maxlev = 4
    if ctl.tag like "*ContrastBorder*" then
        ctl.bordercolor = colorlevels(maxlev)
    else
        ctl.bordercolor = backcol
    end if
    
    '--now, find the control type and apply the applicable
    select case ctl.controltype
        '---
        '---command button
        case accommandbutton, actogglebutton
            ctl.backcolor = backcol
            
            '---this is for swapping out button icons for light / dark theme icons - turned off by default---
            '            if (ctl.picture = "") then goto skipahead0
            '            if darkmode then
            '                if instr(ctl.picture, "\Core_theme_light\") then ctl.picture = replace(ctl.picture, "\Core_theme_light\", "\Core\")
            '            else
            '                if instr(ctl.picture, "\Core\") then ctl.picture = replace(ctl.picture, "\Core\", "\Core_theme_light\")
            '            end if
            '---
            
            
            '---test for individual attributes---
            
            if ctl.tag like "*dis*" then
                fadefore = shadecolor(forebase, disfore)
                ctl.forecolor = fadefore
                ctl.hoverforecolor = fadefore
                ctl.pressedforecolor = fadefore
            else
                fadefore = shadecolor(forebase, levfore - 0.2)
                ctl.forecolor = forebase
                ctl.hoverforecolor = forebase
                ctl.pressedforecolor = forebase
            end if
            
            if ctl.tag like "*btnX*" then
                fadeback = shadecolor(btnxback, scalarback)
                btnxbackshade = shadecolor(btnxback, (0.1 * level) + scalarback)
                ctl.backcolor = btnxbackshade
                ctl.bordercolor = btnxback
            else
                fadeback = shadecolor(backcol, scalarback)
            end if
            
            if ctl.tag like "*accentBtn*" then
                fadeback = shadecolor(backaccent, (0.2 * level) + scalarback)
                ctl.backcolor = shadecolor(backaccent, scalarback)
                ctl.gradient = 17
            end if
            
            ctl.hovercolor = fadeback
            ctl.pressedcolor = fadeback
            
            if ctl.tag like "*cardBtn*" then
                ctl.hovercolor = backcol
                ctl.pressedcolor = backcol
            end if
        '---
        '---label
        case aclabel
            ctl.forecolor = shadecolor(forebase, levfore)
            if ctl.tag like "*lbl_wBack.L#*" then ctl.backcolor = backcol
        '---
        '---text box
        case actextbox, accombobox
            ctl.backcolor = backcol
            if ctl.tag like "*txtTransFore*" then
                ctl.forecolor = backcol
            elseif ctl.tag like "*txtErr*" then
                ctl.bordercolor = btnxback
                ctl.borderstyle = 1
                ctl.forecolor = forebase
            else
                ctl.forecolor = forebase
            end if
            
            if ctl.formatconditions.count = 1 then 'special case for null value conditional formatting. typically this is used for placeholder values
                if ctl.formatconditions.item(0).expression1 like "*IsNull*" then
                    ctl.formatconditions.item(0).backcolor = backcol
                    ctl.formatconditions.item(0).forecolor = forebase
                end if
            end if
        '---
        '---box / subform
        case acrectangle, acsubform
            if not ctl.name like "sfrm*" then ctl.backcolor = backcol
        '---
        '---tab control
        case actabctl
            ctl.pressedcolor = backcol
            fadeback = shadecolor(clng(colorlevels(level - 1)), scalarback)
            ctl.hovercolor = fadeback
            ctl.hoverforecolor = forebase
            ctl.pressedforecolor = forebase
            if level = 0 then
                ctl.backcolor = colorlevels(level + 0)
                fadefore = shadecolor(forebase, levfore - 0.6)
                ctl.forecolor = fadefore
            else
                ctl.backcolor = colorlevels(level - 1)
                fadefore = shadecolor(forebase, levfore)
                ctl.forecolor = fadefore
            end if
        '---
        '---picture
        case acimage
            ctl.backcolor = backcol
    end select
    
nextcontrol:
next

end function

function findcolorlevel(tagtext as string) as long

findcolorlevel = 0
if tagtext = "" then exit function

findcolorlevel = mid(tagtext, instr(tagtext, ".L") + 2, 1)

end function

function shadecolor(inputcolor as long, scalar as double) as long

dim temphex, ior, iog, iob

temphex = hex(inputcolor)

if temphex = "0" then temphex = "111111"

if len(temphex) = 1 then temphex = "0" & temphex
if len(temphex) = 2 then temphex = "0" & temphex
if len(temphex) = 3 then temphex = "0" & temphex
if len(temphex) = 4 then temphex = "0" & temphex
if len(temphex) = 5 then temphex = "0" & temphex

ior = val("&H" & mid(temphex, 5, 2)) * scalar
iog = val("&H" & mid(temphex, 3, 2)) * scalar
iob = val("&H" & mid(temphex, 1, 2)) * scalar

'debug.print ior & " "; iog & " " & iob

if ior > 255 then ior = 255
if iog > 255 then iog = 255
if iob > 255 then iob = 255

if ior < 0 then ior = 0
if iog < 0 then iog = 0
if iob < 0 then iob = 0

shadecolor = rgb(ior, iog, iob)

end function

public function colorpicker(optional lngcolor as long) as long

    'static lngcolor as long
    choosecolor application.hwndaccessapp, lngcolor
    colorpicker = lngcolor

end function
