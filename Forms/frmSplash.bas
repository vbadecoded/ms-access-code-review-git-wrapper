attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()

tempvars.add "loadAmount", 0
tempvars.add "loadWd", 8160

me.lblfrozen.visible = false
call setsplashloading("Loading.")

doevents
form_frmsplash.setfocus
doevents

call setsplashloading("Loading..")

if commandbars("Ribbon").height > 100 then commandbars.executemso "MinimizeRibbon"

call setsplashloading("Loading...")

relinksqltables

'set up theme
dim themeid as long
themeid = nz(dlookup("themeId", "tblParameters"), 0)

dim db as database
set db = currentdb()

dim rstheme as recordset

if themeid <> 0 then
    set rstheme = db.openrecordset("SELECT * FROM tblTheme WHERE recordId = " & themeid)
    
    if rstheme!darkmode.value then
        tempvars.add "themeMode", "Dark"
    else
        tempvars.add "themeMode", "Light"
    end if
    
    tempvars.add "themePrimary", cstr(rstheme!primarycolor.value)
    tempvars.add "themeSecondary", cstr(rstheme!secondarycolor.value)
    tempvars.add "themeAccent", cstr(rstheme!accentcolor.value)
    tempvars.add "themeColorLevels", cstr(rstheme!colorlevels.value)
    
    rstheme.close
    set rstheme = nothing
end if

set db = nothing

call setsplashloading("Loading....")

docmd.openform "_MAIN"
form__main.visible = false

call setsplashloading("Loading.....")

docmd.close acform, "frmSplash"
doevents
form__main.visible = true
docmd.maximize
    
end sub
