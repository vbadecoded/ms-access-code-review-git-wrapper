attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

function applythemechanges()

'all the theme information is in tempvars so it resets when you close it and it will persist an entire database session. this could be a local session variables table as well
tempvars.add "themePrimary", me.primarycolor.value
tempvars.add "themeSecondary", me.secondarycolor.value
tempvars.add "themeAccent", me.accentcolor.value

if me.darkmode then
    tempvars.add "themeMode", "Dark"
else
    tempvars.add "themeMode", "Light"
end if

tempvars.add "themeColorLevels", me.colorlevels.value

'trying to prevent flashing...
docmd.hourglass true
me.painting = false
docmd.echo false

'this code applies the theme to all open forms

dim f as form, sform as control
dim i as integer

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

call settheme(me)
call settheme(me.sfrmthemeeditor.form)

me.showprimary.backcolor = me.primarycolor
me.showsecondary.backcolor = me.secondarycolor
me.showaccent.backcolor = me.accentcolor

'make sure the form updates again
docmd.hourglass false
me.painting = true
docmd.echo true

end function

private sub accentcolor_click()

if me.dirty then me.dirty = false
me.activecontrol = colorpicker(me.activecontrol)

'me.showprimary.backcolor = me.primarycolor
'me.showsecondary.backcolor = me.secondarycolor
me.showaccent.backcolor = me.accentcolor

applythemechanges

end sub

private sub colorlevels_afterupdate()

splitcolorarray

end sub

private sub detail_paint()
on error resume next

me.showprimary.backcolor = me.primarycolor
me.showsecondary.backcolor = me.secondarycolor

end sub

private sub form_load()

call settheme(me)

splitcolorarray
    
end sub

function applylevels()

select case ""
    case nz(me.l1), nz(me.l2), nz(me.l3), nz(me.l4)
        exit function
    case else
        me.colorlevels = me.l1 & "," & me.l2 & "," & me.l3 & "," & me.l4
        applythemechanges
end select

end function

public function splitcolorarray()

dim splitit() as string

splitit = split(me.colorlevels, ",")

me.l1 = splitit(0)
me.l2 = splitit(1)
me.l3 = splitit(2)
me.l4 = splitit(3)

end function

private sub l1_afterupdate()

applylevels

end sub

private sub l2_afterupdate()

applylevels

end sub

private sub l3_afterupdate()

applylevels

end sub

private sub l4_afterupdate()

applylevels

end sub

private sub newtheme_click()

docmd.gotorecord , , acnewrec

end sub

private sub primarycolor_click()

if me.dirty then me.dirty = false
me.activecontrol = colorpicker(me.activecontrol)

me.showprimary.backcolor = me.primarycolor
me.showsecondary.backcolor = me.secondarycolor

applythemechanges

end sub

private sub secondarycolor_click()

if me.dirty then me.dirty = false
me.activecontrol = colorpicker(me.activecontrol)

me.showprimary.backcolor = me.primarycolor
me.showsecondary.backcolor = me.secondarycolor

applythemechanges

end sub

private sub testtheme_click()

if me.dirty then me.dirty = false
applythemechanges

end sub
