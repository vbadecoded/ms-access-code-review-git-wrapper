attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

dim fso as object

private sub cleananddecompose_click()
formstatus (true)

if not cleandatabase then exit sub
doevents

' decompose handles shift bypass internally   no need for user to hold shift
call decomposeaccdb(form__main.cmdrepo & me.cmdrepo.column(2), form__main.cmdrepo)

formstatus (false)
end sub

private sub clearlog_click()
formstatus (true)

dim db as database
set db = currentdb()

db.execute "DELETE * from tblReleaseTracking"

me.sfrmtracking.requery

set db = nothing

formstatus (false)
end sub

private sub cmdrepo_afterupdate()
formstatus (true)

call getrepoinfo(me.cmdrepo)
call gitstatus_click

formstatus (false)
end sub

private sub createaccde_click()
formstatus (true)

dim filepath as string
dim oldname as string, newname as string

filepath = form__main.cmdrepo
oldname = me.cmdrepo.column(2)
newname = left(oldname, len(oldname) - 1) & "e"

dim oaccess as object

set oaccess = createobject("Access.Application")
oaccess.automationsecurity = 1
oaccess.syscmd 603, filepath & oldname, filepath & newname
set oaccess = nothing

addnote "Compiled Version Created: " & filepath & newname

formstatus (false)
end sub

private sub decompose_click()
formstatus (true)

call decomposeaccdb(form__main.cmdrepo & me.cmdrepo.column(2), form__main.cmdrepo)

formstatus (false)
end sub

private sub disableshift_click()
formstatus (true)

if shiftkeybypass(getdb, false) then addnote me.cmdrepo.column(2) & " Shift Key Disabled"

formstatus (false)
end sub

private sub enableshift_click()
formstatus (true)

if shiftkeybypass(getdb, true) then addnote me.cmdrepo.column(2) & " Shift Key Enabled"

formstatus (false)
end sub

private sub form_load()
formstatus (true)

'initial data based on environment variables
me.responsibleperson = environ("username")
me.useremail = getemail(environ("username"))

'----------------------------
'---repository searching
'----------------------------

set fso = createobject("Scripting.FileSystemObject")

dim db as database
set db = currentdb()

dim rsrepos as recordset, rsfindrepo as recordset
set rsrepos = db.openrecordset("tblRepoLocation")

'first delete all records in rsrepo
do while not rsrepos.eof
    rsrepos.delete
    rsrepos.movenext
loop
rsrepos.movefirst

'use fso to scan folders near this repository - these are treated as repositories to work on if an .accdb or .mdb file is found
dim f, sf, sfo
set f = fso.getfolder(fso.getparentfoldername(currentproject.path))
set sf = f.subfolders

dim fsdb, fsdbname as string, fsprodlocname as string
for each sfo in sf
    'look for the record first - skip if found
    if rsrepos.recordcount > 0 then
        set rsfindrepo = db.openrecordset("SELECT * FROM tblRepoLocation WHERE repoLocation = '" & sfo.path & "\" & "'")
        if rsfindrepo.recordcount > 0 then goto skiprepo
    end if
    
    'now scan for the .accdb/.mdb and skip if not found
    fsdbname = ""
    fsprodlocname = ""
    for each fsdb in sfo.files
        select case fso.getextensionname(fsdb.path)
            case "accdb", "mdb"
                'database found!
                fsdbname = fsdb.name
            case "txt"
                if fsdb.name = ".productionLocation.txt" then
                    'get first line of text document
                    open fsdb.path for input as #1
                    line input #1, fsprodlocname
                    close #1
                end if
        end select
    next fsdb
    
    if fsdbname = "" then goto skiprepo 'no database found
    
    rsrepos.addnew
    rsrepos!repolocation = sfo.path & "\"
    rsrepos!dbname = fsdbname
    rsrepos!productionlocation = fsprodlocname
    rsrepos.update
    
skiprepo:
next

'----------------------------
'---repository selection
'----------------------------

'---auto select repo if only one---
if not rsrepos.eof then
    me.cmdrepo = rsrepos!repolocation & "\"
    call getrepoinfo(rsrepos!repolocation & "\")
else
'---if more than one repo found, check tbllastused---
    'if more than one is found, check if the previously used repo is available in the lastused table
    dim rslu as recordset
    set rslu = db.openrecordset("SELECT * from tblLastUsed WHERE recordId = 1")
    if nz(rslu!repolocation, "") = "" then goto lunotfound 'blank field
    set rsfindrepo = db.openrecordset("SELECT * FROM tblRepoLocation WHERE repoLocation = '" & rslu!repolocation & "'")
    if rsfindrepo.recordcount = 1 then 'last used repo found!!
        me.cmdrepo = rslu!repolocation
        call getrepoinfo(rslu!repolocation)
    end if
    
lunotfound:
end if

'---cleanup---
on error resume next
rslu.close: set rslu = nothing
rsfindrepo.close: set rsfindrepo = nothing
rsrepos.close: set rsrepos = nothing
set db = nothing

set fso = nothing


dim dbinit as database
set dbinit = currentdb()
dbinit.execute "DELETE * FROM tblReleaseTracking", dbfailonerror
dbinit.execute "DELETE * FROM tblFiles", dbfailonerror
dbinit.execute "DELETE * FROM tblDiff", dbfailonerror
dbinit.execute "INSERT INTO tblReleaseTracking(task) VALUES('Form Initialized')", dbfailonerror
set dbinit = nothing
me.sfrmtracking.requery

call me.gitstatus_click

formstatus (false)
end sub

private sub gitbranch_afterupdate()
formstatus (true)

addnote "git checkout " & me.gitbranch

call rungitcmd("git checkout " & me.gitbranch)
call gitstatus_click

formstatus (false)
end sub

private sub gitcommit_click()
formstatus (true)

addnote "git commit -m """ & me.releasenotes & """"
call rungitcmd("git commit -m """ & me.releasenotes & """")
call me.gitstatus_click

formstatus (false)
end sub

private sub gitmerge_click()
formstatus (true)

addnote "git merge " & me.gitbranchselect
call rungitcmd("git merge " & me.gitbranchselect)

formstatus (false)
end sub

private sub gitpull_click()
formstatus (true)

addnote "git pull origin " & me.gitbranch
call rungitcmd("git pull origin " & me.gitbranch)

formstatus (false)
end sub

private sub gitpush_click()
formstatus (true)

addnote "git push origin " & me.gitbranch
call rungitcmd("git push origin " & me.gitbranch)

formstatus (false)
end sub

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
    if instr(item, "modified:") then
        rsfiles.addnew
        rsfiles!location = trim(replace(item, "modified:", ""))
        rsfiles!filestatus = itemstatus
        rsfiles.update
    elseif itemstatus = "new" then
        if fso.fileexists(me.cmdrepo & replace(item, chr(9), "")) then
            rsfiles.addnew
            rsfiles!location = trim(replace(item, "modified:", ""))
            rsfiles!filestatus = itemstatus
            rsfiles.update
        end if
    end if
next item

rsfiles.close
set rsfiles = nothing
set dbstatus = nothing

set fso = nothing

me.sfrmfiles.requery
me.sfrmdiff.requery

formstatus (false)
end sub

private sub increaserev_click()
formstatus (true)

dim x, y, major, minor, min, newmajor, newminor, newmin

x = me.releasenum
y = replace(x, "REV", "")
major = split(y, ".")(0)
minor = split(y, ".")(1)
min = split(y, ".")(2)

if (min <> 99) then
    newmajor = major
    newminor = minor
    newmin = min + 1
    if newmin < 10 then newmin = "0" & newmin
    goto done
end if
newmin = "00"

if (minor <> 9) then
    newmajor = major
    newminor = minor + 1
    goto done
end if
newminor = 0
newmajor = major + 1

done:
dim newrel as string
newrel = "REV" & newmajor & "." & newminor & "." & newmin
me.releasenum = newrel
addnote "Rev Increased to " & newrel

formstatus (false)
end sub

function formstatus(inwork as boolean)

if inwork then
    me.detail.backcolor = rgb(50, 0, 0)
else
    call settheme(me)
end if

me.coderunning.visible = inwork

end function

private sub notifydepartment_afterupdate()
formstatus (true)

dim db as database
dim rs as recordset

set db = opendatabase("\\data\mdbdata\WorkingDB\_docs\Reporting\WorkingDB_ForExcel.accdb", , true)
set rs = db.openrecordset("SELECT * FROM tblPermissions WHERE inactive = false AND dept = '" & me.notifydepartment & "'")

dim emails as string
emails = ""

do while not rs.eof
    emails = emails & rs!useremail & "; "
    rs.movenext
loop

call genemail(strbcc:=emails, strsubject:="WorkingDB Update Released", body:=me.releasenotes)

rs.close
set rs = nothing
set db = nothing

addnote me.notifydepartment & " email generated"

formstatus (false)
end sub

private sub notifyuser_afterupdate()
formstatus (true)

dim db as database
dim rs as recordset

set db = opendatabase("\\data\mdbdata\WorkingDB\_docs\Reporting\WorkingDB_ForExcel.accdb", , true)
set rs = db.openrecordset("SELECT * FROM tblPermissions WHERE user = '" & me.notifyuser & "'")

call genemail(strto:=rs!useremail, strsubject:="WorkingDB Update Released", body:=me.releasenotes)

rs.close
set rs = nothing
set db = nothing

addnote me.notifyuser & " email generated"

formstatus (false)
end sub

private sub openaccdb_click()
formstatus (true)

call openpath(getdb)
addnote me.cmdrepo.column(2) & " Opened"

formstatus (false)
end sub

private sub opengitgui_click()
formstatus (true)

addnote "git gui"
call rungitcmd("git gui")

formstatus (false)
end sub

private sub openthemeeditor_click()
formstatus (true)

addnote "open theme editor"

docmd.openform "frmThemeEditor"

formstatus (false)
end sub

private sub publishchanges_click()
formstatus (true)

addnote "git pull origin master : " & me.cmdrepo.column(1)
call rungitcmd("git pull origin master", me.cmdrepo.column(1))

formstatus (false)
end sub

private sub publishfe_click()
formstatus (true)

call cleandatabase

formstatus (false)
end sub

private sub publishnotes_click()
formstatus (true)

currentdb.execute "INSERT INTO " & me.revisiontablename & _
    "(DatabaseVersion,Notes,ReleaseDate,ReleasedBy,DatabaseName)" & _
    " VALUES" & _
    "('" & me.releasenum & "','" & me.releasenotes & "','" & date & "','" & me.responsibleperson & "','" & me.cmdrepo.column(2) & "');", dbfailonerror

dim body, strvalues
addnote "Generate notification email"
body = emailcontentgen("New Version Published", me.cmdrepo.column(2) & " " & me.releasenum & " Published", "Notes: " & replace(me.releasenotes, ",", ";"), "Responsible: " & responsibleperson, "Releaser: " & environ("username"), "", "")

if environ("username") <> "brownj" then
    strvalues = "'brownj','brownj@us.nifco.com','" & environ("username") & "','" & getemail(environ("username")) & "','" & now() & "',1,1,'New Version Published','" & body & "','" & now() & "'"
    currentdb.execute "INSERT INTO tblNotificationsSP(recipientUser,recipientEmail,senderUser,senderEmail,sentDate,notificationType,notificationPriority,notificationDescription,emailContent,readDate) VALUES(" & strvalues & ");", dbfailonerror
    addnote "Notification sent to brownj"
end if

if environ("username") <> "georgemi" then
    strvalues = "'georgemi','georgemi@us.nifco.com','" & environ("username") & "','" & getemail(environ("username")) & "','" & now() & "',1,1,'New Version Published','" & body & "','" & now() & "'"
    currentdb.execute "INSERT INTO tblNotificationsSP(recipientUser,recipientEmail,senderUser,senderEmail,sentDate,notificationType,notificationPriority,notificationDescription,emailContent,readDate) VALUES(" & strvalues & ");", dbfailonerror
    addnote "Notification sent to georgemi"
end if

addnote "Version " & me.releasenum & " Notes Published Successfully"

formstatus (false)
end sub

private sub recomposesendfile_click()
formstatus (true)
set fso = createobject("Scripting.FileSystemObject")

'---add all changes files to list---
dim results as string
results = rungitcmd("git status")

dim dbrecomp as database
set dbrecomp = currentdb()
dbrecomp.execute "DELETE * FROM tblFiles", dbfailonerror

dim arr() as string
arr = split(results, vblf)

dim item
dim rsrecomp as dao.recordset
set rsrecomp = dbrecomp.openrecordset("tblFiles", dbopendynaset, dbappendonly)

for each item in arr
    if instr(item, "modified:") then
        rsrecomp.addnew
        rsrecomp!location = trim(replace(item, "modified:", "")) & " "
        rsrecomp.update
    end if
next item

rsrecomp.close
set rsrecomp = nothing
set dbrecomp = nothing

formstatus (false)
end sub

private sub releasehelp_click()
formstatus (true)

followhyperlink "https://github.com/workingdb/workingdb?tab=contributing-ov-file"
addnote "Opened Help Page"

formstatus (false)
end sub

private sub responsibleperson_afterupdate()
formstatus (true)

if me.dirty then me.dirty = false
me.useremail = getemail(me.responsibleperson)
addnote "Populated User Email"

formstatus (false)
end sub

private sub stagechanged_click()
formstatus (true)
addnote "git add ."

call rungitcmd("git add .")
call gitstatus_click

formstatus (false)
end sub

private sub trackrevisions_click()
formstatus (true)

dim vis as boolean
vis = me.trackrevisions

me.label196.visible = vis
me.revisiontablename.visible = vis
me.publishnotes.visible = vis
me.releasenum.visible = vis
me.label67.visible = vis
me.command76.visible = vis
me.increaserev.visible = vis
me.responsibleperson.visible = vis
me.lblresp.visible = vis
me.respbackg.visible = vis
me.useremail.visible = vis

formstatus (false)
end sub
