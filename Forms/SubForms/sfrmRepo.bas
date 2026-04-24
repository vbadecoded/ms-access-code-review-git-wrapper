attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

public sub gitstatus_click()
formstatus (true)

addnote "git status"

'add all modified files
dim results as string
results = rungitcmd("git status", printall:=false)

dim dbstatus as database
set dbstatus = currentdb()
dbstatus.execute "DELETE * FROM tblFiles", dbfailonerror
dbstatus.execute "DELETE * FROM tblDiff", dbfailonerror

dim fso as object
set fso = createobject("Scripting.FileSystemObject")

dim arr() as string
arr = split(results, vblf)

dim item, itemstatus as string
dim rsfiles as dao.recordset
set rsfiles = dbstatus.openrecordset("tblFiles", dbopendynaset, dbappendonly)

for each item in arr
    if instr(item, "Changes to be committed") then itemstatus = "staged"
    if instr(item, "Changes not staged for commit") then itemstatus = "unstaged"
    if instr(item, "Untracked files") then itemstatus = "new"
    if instr(item, "deleted") then itemstatus = "deleted"
    if instr(item, "modified:") then
        rsfiles.addnew
        rsfiles!location = trim(replace(replace(item, "modified:", ""), chr(9), ""))
        rsfiles!filestatus = itemstatus
        rsfiles.update
    elseif itemstatus = "new" then
        if fso.fileexists(form_sfrmrepo.cmdrepo & replace(item, chr(9), "")) then
            rsfiles.addnew
            rsfiles!location = replace(replace(item, "modified:", ""), chr(9), "")
            rsfiles!filestatus = itemstatus
            rsfiles.update
        end if
    elseif itemstatus = "deleted" then
        if len(replace(replace(item, "deleted:", ""), chr(9), "")) > 1 then
            rsfiles.addnew
            rsfiles!location = replace(replace(item, "deleted:", ""), chr(9), "")
            rsfiles!filestatus = itemstatus
            rsfiles.update
        end if
    end if
next item

rsfiles.close
set rsfiles = nothing
set dbstatus = nothing

set fso = nothing

form__main.sfrmfiles.requery
form__main.sfrmfiles.requery

call formstatus(false)
end sub

private sub stagechanged_click()
formstatus (true)
addnote "git add ."

call rungitcmd("git add .")
call gitstatus_click

formstatus (false)
end sub

private sub gitcommit_click()
formstatus (true)

addnote "git commit -m """ & form_sfrmrepo.releasenotes & """"
call rungitcmd("git commit -m """ & form_sfrmrepo.releasenotes & """")
call me.gitstatus_click

formstatus (false)
end sub

private sub gitpush_click()
formstatus (true)

addnote "git push origin " & me.gitbranch
call rungitcmd("git push origin " & me.gitbranch)

formstatus (false)
end sub

private sub opengitgui_click()
formstatus (true)

addnote "git gui"
call rungitcmd("git gui")

formstatus (false)
end sub

private sub publishchanges_click()
formstatus (true)

addnote "git pull origin master : " & form_sfrmrepo.cmdrepo.column(1)
call rungitcmd("git pull origin master", form_sfrmrepo.cmdrepo.column(1))

formstatus (false)
end sub

private sub cmdrepo_afterupdate()
formstatus (true)

call getrepoinfo(form_sfrmrepo.cmdrepo)
call gitstatus_click

me.filter = "repoLocation = '" & me.activecontrol & "'"
me.filteron = true

if nz(me.repolocation, "") = "" then me.repolocation = me.activecontrol

formstatus (false)
end sub

private sub gitpull_click()
formstatus (true)

addnote "git pull origin " & me.gitbranch
call rungitcmd("git pull origin " & me.gitbranch)

formstatus (false)
end sub

private sub gitmerge_click()
formstatus (true)

addnote "git merge " & me.gitbranchselect
call rungitcmd("git merge " & me.gitbranchselect)

formstatus (false)
end sub

private sub gitbranch_afterupdate()
formstatus (true)

addnote "git checkout " & me.gitbranch

call rungitcmd("git checkout " & me.gitbranch)
call gitstatus_click

formstatus (false)
end sub
