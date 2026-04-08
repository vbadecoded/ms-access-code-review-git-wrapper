attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub location_click()

dim gitcmd as string

if me.filestatus = "staged" then
    gitcmd = "git diff --cached "
else
    gitcmd = "git diff "
end if

addnote gitcmd & me.location

'add all modified files
dim results as string
results = rungitcmd(gitcmd & trim(me.location), printnone:=true)

dim dbdiff as database
set dbdiff = currentdb()
dbdiff.execute "DELETE * FROM tblDiff", dbfailonerror

dim arr() as string
arr = split(results, vblf)

dim item
dim rsdiff as dao.recordset
set rsdiff = dbdiff.openrecordset("tblDiff", dbopendynaset, dbappendonly)

for each item in arr
    rsdiff.addnew
    rsdiff!diffline = item
    rsdiff.update
next item

rsdiff.close
set rsdiff = nothing
set dbdiff = nothing

form__main.lblgitdiff.caption = "Git Diff " & me.location
form__main.sfrmdiff.requery
end sub

private sub stage_click()

if me.filestatus <> "staged" then
    call rungitcmd("git add " & me.location)
    call form__main.gitstatus_click
else
    call rungitcmd("git reset " & me.location)
    call form__main.gitstatus_click
end if


end sub
