attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub detail_paint()
on error resume next

me.primarycolor.backcolor = me.primarycolor
me.primarycolor.forecolor = me.primarycolor

if me.secondarycolor = 0 then
    me.secondarycolor.backcolor = me.primarycolor
    me.secondarycolor.forecolor = me.primarycolor
else
    me.secondarycolor.backcolor = me.secondarycolor
    me.secondarycolor.forecolor = me.secondarycolor
end if

if me.darkmode then
    me.dmode.backcolor = 0
    me.dmode.forecolor = vbwhite
    me.themename.forecolor = vbwhite
else
    me.dmode.backcolor = vbwhite
    me.dmode.forecolor = 0
    me.themename.forecolor = 0
end if

end sub

private sub form_load()

call settheme(me)
    
end sub

private sub themename_click()

dim f as form, sform as control
dim i as integer

tempvars.add "themePrimary", me.primarycolor.value
tempvars.add "themeSecondary", me.secondarycolor.value
tempvars.add "themeAccent", me.accentcolor.value

if me.darkmode then
    tempvars.add "themeMode", "Dark"
else
    tempvars.add "themeMode", "Light"
end if

tempvars.add "themeColorLevels", me.colorlevels.value

form_frmthemeeditor.filter = "recordId = " & me.recordid
form_frmthemeeditor.filteron = true

call assignthemetoparameters(me.recordid)

form_frmthemeeditor.splitcolorarray
doevents

dim obj

for each obj in application.currentproject.allforms
    if obj.isloaded = false then goto nextone
    set f = forms(obj.name)
    call settheme(f)
    for each sform in f.controls
        if sform.controltype = acsubform then
            on error resume next
            call settheme(sform.form)
        end if
    next sform
nextone:
next obj

end sub
