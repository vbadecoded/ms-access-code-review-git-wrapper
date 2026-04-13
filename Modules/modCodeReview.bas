option compare database
option explicit

private declare ptrsafe sub sleep lib "kernel32" (byval dwmilliseconds as long)

function cleandatabase() as boolean

cleandatabase = false

' --- validation ---
dim errormsg as new collection
if nz(form__main.releasenotes, "") = "" then errormsg.add "Empty release notes"

if errormsg.count > 0 then
    dim msgcontents as string, item
    for each item in errormsg
        msgcontents = msgcontents & vbnewline & item
    next item
    msgbox msgcontents, vbinformation, "Please fix these issues: "
    exit function
end if

if msgbox("Are you sure? ", vbyesno, "Just Checking") = vbno then exit function

addnote "--- STARTING " & form__main.cmdrepo.column(2) & " CLEANING PROCEDURE ---"

' --- setup variables (safe add — remove first if exists) ---
addnote "Establishing variables..."

dim fso as object
set fso = createobject("Scripting.FileSystemObject")

dim repoloc as string, devfile as string, dbname as string
repoloc = form__main.cmdrepo
devfile = getdb
dbname = form__main.cmdrepo.column(2)

safeaddtempvar "releaseNum", form__main.releasenum.value
safeaddtempvar "releaseNotes", replace(form__main.releasenotes.value, "'", "''")
safeaddtempvar "responsiblePerson", form__main.responsibleperson.value
safeaddtempvar "userEmail", form__main.useremail.value
safeaddtempvar "databaseName", dbname
safeaddtempvar "devFile", devfile

' --- front-end specific cleanup ---
if dbname = "WorkingDB_FE.accdb" then
    addnote "Opening database for cleaning/compiling"
    dim dbinputrs as database
    
    set dbinputrs = opendatabase(devfile)
    dbinputrs.execute "DELETE FROM tblPLM"
    dbinputrs.execute "DELETE * FROM tblSessionVariables"
    dbinputrs.execute "UPDATE [tblDBinfo] SET [Release] = '" & tempvars!releasenum & "' WHERE [ID] = 1"
    dbinputrs.close
    set dbinputrs = nothing
    
    ' run readyforpublish check — reuse this instance for decompile too
    addnote "Running readyForPublish check..."
    dim dbinput as object
    set dbinput = createobject("Access.Application")
    dbinput.opencurrentdatabase devfile
    dbinput.runcommand accmdcloseall
    
    do
    loop until dbinput.run("readyForPublish") = true
    
    dbinput.closecurrentdatabase
    dbinput.quit
    set dbinput = nothing
    
    ' backup backends
    dim dbloc as string, bebackup as string
    dbloc = "\\data\mdbdata\WorkingDB\"
    bebackup = dbloc & "_backups\prod-BE\"
    
    addnote "Backup backends"
    dim befiles as variant
    befiles = array("WorkingDB_BE.accdb", "WorkingDB_BE_DesignE.accdb", "WorkingDB_BE_ProjectE.accdb")
    
    dim i as long
    for i = lbound(befiles) to ubound(befiles)
        fso.copyfile dbloc & "prod-BE\" & befiles(i), bebackup & tempvars!releasenum & "_" & befiles(i)
    next i
end if

' --- enable shift bypass + decompile in one access instance ---
addnote "Enable Shift Bypass + Decompile"
dim accdecompile as object
set accdecompile = createobject("Access.Application")

' set bypass via dbengine (no startup code)
accdecompile.dbengine.opendatabase(devfile, false, false).properties("AllowByPassKey") = true

' now decompile using the same instance — /decompile flag compiles then removes compiled state
addnote "Decompiling..."
accdecompile.opencurrentdatabase devfile, false  ' false = don't execute startup
accdecompile.runcommand accmdcloseall
accdecompile.closecurrentdatabase
accdecompile.quit
set accdecompile = nothing

' --- compact / repair dev into temp ---
addnote "Compacting dev into temp file"
dim devtemp as string
devtemp = repoloc & "temp.accdb"
application.compactrepair devfile, devtemp
fso.deletefile devfile

' --- compile temp file ---
' reuse a single access instance for compile + backup
addnote "Compile"
call shiftkeybypass(devtemp, true)

dim dbtemp as object
set dbtemp = createobject("Access.Application")
dbtemp.opencurrentdatabase devtemp
dbtemp.visible = false
dbtemp.runcommand accmdcloseall

dim compileme as object
set compileme = dbtemp.vbe.commandbars.findcontrol(msocontrolbutton, 578)
if compileme.enabled then compileme.execute

dbtemp.runcommand accmdcompileandsaveallmodules

' backup to home drive
dim backupdir as string
backupdir = "H:\wdbBackups\"
if not fso.folderexists(backupdir) then mkdir backupdir
addnote "Backup temp into homedrive"
fso.copyfile devtemp, backupdir & "WorkingDB_Dev_backup.accdb"

' disable shift bypass and close
addnote "Disable shift bypass"
dbtemp.currentdb.properties("AllowByPassKey") = false
addnote "Close temp file"
dbtemp.closecurrentdatabase
dbtemp.quit
set dbtemp = nothing
doevents

' --- compact temp back into dev ---
addnote "Compacting temp file back into FE"
application.compactrepair devtemp, devfile
fso.deletefile devtemp

set fso = nothing
addnote dbname & " CLEANED"

cleandatabase = true

end function

private sub safeaddtempvar(varname as string, varvalue as variant)
    on error resume next
    tempvars.remove varname
    on error goto 0
    tempvars.add varname, varvalue
end sub

function getrepoinfo(repolocation as string) as boolean
getrepoinfo = false

if form__main.trackrevisions then
    addnote "Getting " & form__main.cmdrepo & " latest revision"
    
    dim maxrel as variant
    maxrel = dmax("ID", form__main.revisiontablename, "databaseName = '" & form__main.cmdrepo.column(2) & "'")
    
    form__main.releasenum = dlookup("DatabaseVersion", form__main.revisiontablename, "ID = " & nz(maxrel, 0))
end if

'find current branch
form__main.gitbranch = rungitcmd("git branch --show-current")

'list branches
form__main.gitbranch.rowsource = replace(replace(rungitcmd("git branch"), vblf, ";"), "*", "")
form__main.gitbranchselect.rowsource = form__main.gitbranch.rowsource

form__main.publishchanges.visible = nz(form__main.cmdrepo.column(1), "") <> ""

currentdb.execute "UPDATE tblLastUsed SET repoLocation = '" & repolocation & "' WHERE recordId = 1", dbfailonerror

getrepoinfo = true
end function

function getdb() as string
getdb = form__main.cmdrepo & form__main.cmdrepo.column(2)
end function

function shiftkeybypass(location as string, toggle as boolean) as boolean
shiftkeybypass = false
on error goto errhandler

dim db as dao.database
dim acc as object
dim prop as dao.property
const conpropnotfound = 3270
  
set acc = createobject("Access.Application")
set db = acc.dbengine.opendatabase(location, false, false)

db.properties("AllowByPassKey") = toggle
shiftkeybypass = true

exitthis:
on error resume next
if not (db is nothing) then db.close
set db = nothing
if not (acc is nothing) then acc.quit
set acc = nothing
exit function

errhandler:
if err.number = conpropnotfound then
    set prop = db.createproperty("AllowByPassKey", dbboolean, toggle)
    db.properties.append prop
    shiftkeybypass = true
    resume exitthis
end if

msgbox "Error: Maybe you already have it open? Please close if so." & vbnewline & _
       "Details: " & err.description, vbexclamation, "Shift Bypass Failed"
resume exitthis

end function

function rungitcmd(inputcmd as string, optional dir as string = "current", optional printall as boolean = true, optional printnone as boolean = false) as string

dim wsshell as object
dim sworkingdirectory as string
dim stroutput as string, strerror as string

' set the working directory to your git repository
rungitcmd = ""
if isnull(form__main.cmdrepo) then exit function
if dir = "current" then
    sworkingdirectory = form__main.cmdrepo
else
    sworkingdirectory = dir
end if

set wsshell = createobject("WScript.Shell")
wsshell.currentdirectory = sworkingdirectory

select case inputcmd
    case "git commit -a"
        inputcmd = inputcmd & " -m """ & form__main.releasenotes & """"
end select

' use .run with hidden window (0) and redirect stdout + stderr to temp files
dim tmpout as string, tmperr as string
tmpout = environ("temp") & "\git_stdout.txt"
tmperr = environ("temp") & "\git_stderr.txt"

wsshell.run "cmd /c " & inputcmd & " > """ & tmpout & """ 2> """ & tmperr & """", 0, true

' read stdout
dim fso as object
set fso = createobject("Scripting.FileSystemObject")

stroutput = ""
if fso.fileexists(tmpout) then
    if fso.getfile(tmpout).size > 0 then
        stroutput = fso.opentextfile(tmpout).readall()
    end if
    fso.deletefile tmpout
end if

' read stderr
strerror = ""
if fso.fileexists(tmperr) then
    if fso.getfile(tmperr).size > 0 then
        strerror = fso.opentextfile(tmperr).readall()
    end if
    fso.deletefile tmperr
end if

set fso = nothing

' log any errors from stderr
if len(strerror) > 0 then
    addnote strerror
end if

dim arr() as string
arr = split(stroutput, vblf)

dim item, startchanges as boolean
startchanges = false

if printnone then goto noprint

' use recordset.addnew for bulk inserts — much faster than docmd.runsql per row
dim db as database, rs as dao.recordset
set db = currentdb()
set rs = db.openrecordset("tblReleaseTracking", dbopendynaset, dbappendonly)

for each item in arr
    if item = "" then startchanges = true
    
    if startchanges = true and printall = false then
        rs.close
        goto noprint
    end if
    
    rs.addnew
    rs!task = item & " "
    rs.update
next item

rs.close
set rs = nothing
set db = nothing

noprint:

on error resume next
form_sfrmtracking.requery

movetrackingtolastrecord

rungitcmd = stroutput

set wsshell = nothing

end function

function movetrackingtolastrecord()

on error resume next

' set focus to the subform control so we can navigate it
form__main.sfrmtracking.setfocus

' use recordsetclone to move to the last record
dim rs as dao.recordset
set rs = form_sfrmtracking.recordsetclone

if rs.recordcount > 0 then
    rs.movelast
    form_sfrmtracking.bookmark = rs.bookmark
end if

set rs = nothing

end function

function recomposeaccdb()

dim fso as object
set fso = createobject("Scripting.FileSystemObject")

dim repo as string, dbsource as string, importto as string
repo = form__main.cmdrepo
dbsource = repo & form__main.cmdrepo.column(2)

' create a working copy of the .accdb to import into
importto = repo & "recompose_temp.accdb"
if fso.fileexists(importto) then fso.deletefile importto
addnote "Copying " & form__main.cmdrepo.column(2) & " to temp file..."
fso.copyfile dbsource, importto

' ask user: import all modified files or only selected?
dim importmode as integer
importmode = msgbox("Yes = Import ALL modified files" & vbnewline & _
                     "No = Import only SELECTED files", _
                     vbyesnocancel + vbquestion, "Recompose - Choose Import Mode")

if importmode = vbcancel then
    addnote "Recompose cancelled"
    fso.deletefile importto
    set fso = nothing
    exit function
end if

' build a collection of file paths to import
dim db as database
set db = currentdb()

dim rs as recordset
if importmode = vbyes then
    set rs = db.openrecordset("SELECT location FROM tblFiles WHERE Nz(fileStatus,'') <> ''")
else
    set rs = db.openrecordset("SELECT location FROM tblFiles WHERE selected = True")
end if

if rs.eof then
    addnote "No files to import"
    rs.close: set rs = nothing: set db = nothing
    fso.deletefile importto
    set fso = nothing
    exit function
end if

' collect file paths into a dictionary for quick lookup
dim filestoimport as object
set filestoimport = createobject("Scripting.Dictionary")
do while not rs.eof
    filestoimport.add trim(rs!location), true
    rs.movenext
loop
rs.close: set rs = nothing: set db = nothing

' open the temp database
addnote "Starting Access..."
dim oapplication as object
set oapplication = createobject("Access.Application")
addnote "Opening " & importto & " ..."
oapplication.opencurrentdatabase importto
oapplication.runcommand accmdcloseall
oapplication.currentdb.properties("AllowByPassKey") = true

dim folder as object, myfile as object
dim objectname as string, objecttype as string, relativepath as string

' --- forms ---
if fso.folderexists(repo & "Forms\") then
    set folder = fso.getfolder(repo & "Forms\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Forms/" & myfile.name
        if objecttype = "form" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading form: " & objectname
            oapplication.loadfromtext acform, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- subforms ---
if fso.folderexists(repo & "Forms\SubForms\") then
    set folder = fso.getfolder(repo & "Forms\SubForms\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Forms/SubForms/" & myfile.name
        if objecttype = "form" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading subform: " & objectname
            oapplication.loadfromtext acform, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- modules ---
if fso.folderexists(repo & "Modules\") then
    set folder = fso.getfolder(repo & "Modules\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Modules/" & myfile.name
        if objecttype = "bas" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading module: " & objectname
            oapplication.loadfromtext acmodule, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- macros ---
if fso.folderexists(repo & "Macros\") then
    set folder = fso.getfolder(repo & "Macros\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Macros/" & myfile.name
        if objecttype = "mod" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading macro: " & objectname
            oapplication.loadfromtext acmacro, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- queries ---
if fso.folderexists(repo & "Queries\") then
    set folder = fso.getfolder(repo & "Queries\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Queries/" & myfile.name
        if objecttype = "sql" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading query: " & objectname
            importqueryfromsql oapplication, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- subqueries ---
if fso.folderexists(repo & "Queries\SubQueries\") then
    set folder = fso.getfolder(repo & "Queries\SubQueries\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Queries/SubQueries/" & myfile.name
        if objecttype = "sql" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading subquery: " & objectname
            importqueryfromsql oapplication, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- reports ---
if fso.folderexists(repo & "Reports\") then
    set folder = fso.getfolder(repo & "Reports\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Reports/" & myfile.name
        if objecttype = "rpt" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading report: " & objectname
            oapplication.loadfromtext acreport, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' --- subreports ---
if fso.folderexists(repo & "Reports\SubReports\") then
    set folder = fso.getfolder(repo & "Reports\SubReports\")
    for each myfile in folder.files
        objecttype = fso.getextensionname(myfile.name)
        objectname = fso.getbasename(myfile.name)
        relativepath = "Reports/SubReports/" & myfile.name
        if objecttype = "rpt" and shouldimport(filestoimport, relativepath, importmode) then
            addnote "Loading subreport: " & objectname
            oapplication.loadfromtext acreport, objectname, myfile.path
            addnote objectname & " LOADED"
        end if
    next
end if

' compile and close
oapplication.runcommand accmdcompileandsaveallmodules
oapplication.closecurrentdatabase
oapplication.quit
set oapplication = nothing

' replace original with recomposed version
addnote "Replacing original with recomposed file..."
fso.deletefile dbsource
fso.copyfile importto, dbsource
fso.deletefile importto

set filestoimport = nothing
set fso = nothing

addnote "Recompose Complete"

end function

function shouldimport(filestoimport as object, relativepath as string, importmode as integer) as boolean
    ' in "all modified" or "selected" mode, check if this file is in the list
    ' normalize slashes for comparison since git uses forward slashes
    dim normalized as string
    normalized = replace(relativepath, "\", "/")
    
    dim key as variant
    for each key in filestoimport.keys
        if instr(1, replace(trim(key), "\", "/"), normalized, vbtextcompare) > 0 then
            shouldimport = true
            exit function
        end if
    next key
    
    shouldimport = false
end function

function decomposeaccdb(sadpfilename as string, sexportpath as string)

dim fso as object
set fso = createobject("Scripting.FileSystemObject")

dim mytype as string, myname as string
mytype = fso.getextensionname(sadpfilename)
myname = fso.getbasename(sadpfilename)

' copy to a stub file so we don't modify the original
dim sstubadpfilename as string
sstubadpfilename = environ("temp") & "\" & myname & "_stub." & mytype
addnote "Copying stub to " & sstubadpfilename & "..."
fso.copyfile sadpfilename, sstubadpfilename

' enable shift bypass via dbengine then open in the same instance
addnote "Starting Access..."
dim oapplication as object
set oapplication = createobject("Access.Application")

' set bypass via raw db access first (no startup code runs)
addnote "Enabling shift bypass..."
dim dbbypass as object
set dbbypass = oapplication.dbengine.opendatabase(sstubadpfilename, false, false)
dbbypass.properties("AllowByPassKey") = true
dbbypass.close
set dbbypass = nothing

' now open with full access in the same instance — bypass is already active
oapplication.opencurrentdatabase sstubadpfilename
oapplication.visible = false

addnote "Clearing old export files..."
' clear all export folders in one pass
dim exportfolders as variant
exportfolders = array("Forms\", "Forms\SubForms\", "Modules\", "Macros\", _
                      "Reports\", "Reports\SubReports\", "Queries\", _
                      "Queries\SubQueries\", "Tables\", "VBproject\")

dim i as long
for i = lbound(exportfolders) to ubound(exportfolders)
    clearfolder fso, sexportpath & "\" & exportfolders(i)
next i

' ensure all output folders exist
for i = lbound(exportfolders) to ubound(exportfolders)
    ensurefolder fso, sexportpath & "\" & exportfolders(i)
next i

addnote "Exporting..."
dim myobj as object

' --- forms ---
addnote "       exporting forms"
for each myobj in oapplication.currentproject.allforms
    if left(myobj.fullname, 1) = "s" then
        oapplication.saveastext acform, myobj.fullname, sexportpath & "\Forms\SubForms\" & myobj.fullname & ".form"
        splitformfile (sexportpath & "\Forms\SubForms\" & myobj.fullname & ".form")
        normalizeexportfile sexportpath & "\Forms\SubForms\" & myobj.fullname & ".bas"
    else
        oapplication.saveastext acform, myobj.fullname, sexportpath & "\Forms\" & myobj.fullname & ".form"
        splitformfile (sexportpath & "\Forms\" & myobj.fullname & ".form")
        normalizeexportfile sexportpath & "\Forms\" & myobj.fullname & ".bas"
    end if
next

' --- modules ---
addnote "       exporting modules"
for each myobj in oapplication.currentproject.allmodules
    oapplication.saveastext acmodule, myobj.fullname, sexportpath & "\Modules\" & myobj.fullname & ".bas"
    normalizeexportfile sexportpath & "\Modules\" & myobj.fullname & ".bas"
next

' --- macros ---
addnote "       exporting macros"
for each myobj in oapplication.currentproject.allmacros
    oapplication.saveastext acmacro, myobj.fullname, sexportpath & "\Macros\" & myobj.fullname & ".mod"
    'normalizeexportfile sexportpath & "\Macros\" & myobj.fullname & ".mod"
next
flushnotes

' --- reports ---
addnote "       exporting reports"
for each myobj in oapplication.currentproject.allreports
    if left(myobj.fullname, 1) = "s" then
        oapplication.saveastext acreport, myobj.fullname, sexportpath & "\Reports\SubReports\" & myobj.fullname & ".rpt"
        'normalizeexportfile sexportpath & "\Reports\SubReports\" & myobj.fullname & ".rpt"
    else
        oapplication.saveastext acreport, myobj.fullname, sexportpath & "\Reports\" & myobj.fullname & ".rpt"
        'normalizeexportfile sexportpath & "\Reports\" & myobj.fullname & ".rpt"
    end if
next
flushnotes

' --- queries ---
addnote "       exporting queries"
' write .sql files for all queries (fast, good for diffs).
' also write .qry via saveastext for passthrough queries only (createquerydef can't handle them).
for each myobj in oapplication.currentdb.querydefs
    if left(myobj.name, 3) = "~sq" then goto nextquery  ' skip form-embedded queries
    if left(myobj.name, 1) = "s" then
        writetotextfile sexportpath & "\Queries\SubQueries\" & myobj.name & ".sql", myobj.sql
        if myobj.type = dbqsqlpassthrough then
            oapplication.saveastext acquery, myobj.name, sexportpath & "\Queries\SubQueries\" & myobj.name & ".qry"
        end if
    else
        writetotextfile sexportpath & "\Queries\" & myobj.name & ".sql", myobj.sql
        if myobj.type = dbqsqlpassthrough then
            oapplication.saveastext acquery, myobj.name, sexportpath & "\Queries\" & myobj.name & ".qry"
        end if
    end if
nextquery:
next

' --- vb project information ---
addnote "       exporting vbproject information"
dim dictsubvalues as object, dictbody as object
set dictsubvalues = createobject("Scripting.Dictionary")
set dictbody = createobject("Scripting.Dictionary")

for each myobj in oapplication.vbe.activevbproject.references
    dictsubvalues.add myobj.name & " " & myobj.major & "." & myobj.minor, myobj.fullpath
next

dictbody.add "project-name", oapplication.vbe.activevbproject.name
dictbody.add "vb-references", dictsubvalues
writetotextfile sexportpath & "\VBproject\VBproject-properties.json", tojson(dictbody)

' cleanup
addnote "Cleaning up..."
set myobj = nothing
oapplication.closecurrentdatabase
oapplication.quit
set oapplication = nothing
doevents

' wait for access to fully release the file lock before deleting
dim retries as long
for retries = 1 to 10
    if fso.fileexists(sstubadpfilename) then
        on error resume next
        fso.deletefile sstubadpfilename
        if err.number = 0 then
            on error goto 0
            exit for
        end if
        on error goto 0
    end if
    doevents
    sleep 500  ' wait 500ms between retries
next retries

set fso = nothing

flushnotes
addnote "++ Files Decomposed from " & sadpfilename

end function

private sub clearfolder(fso as object, folderpath as string)
    if not fso.folderexists(folderpath) then exit sub
    dim f as object
    for each f in fso.getfolder(folderpath).files
        fso.deletefile f.path, true
    next
end sub

private sub ensurefolder(fso as object, folderpath as string)
    if fso.folderexists(folderpath) then exit sub
    ' create parent first if needed
    dim parent as string
    parent = fso.getparentfoldername(folderpath)
    if not fso.folderexists(parent) then mkdir parent
    mkdir folderpath
end sub

function writetotextfile(filelocation as string, texttowrite as string)

dim filenum as integer
filenum = freefile()

open filelocation for output as #filenum
print #filenum, texttowrite
close #filenum

end function

sub importqueryfromsql(oapp as object, queryname as string, sqlfilepath as string)
    ' check if a .qry file exists alongside the .sql — if so, it's a passthrough query
    ' and we need loadfromtext to preserve connect/returnsrecords properties
    dim qryfilepath as string
    qryfilepath = replace(sqlfilepath, ".sql", ".qry")
    
    dim fso as object
    set fso = createobject("Scripting.FileSystemObject")
    
    if fso.fileexists(qryfilepath) then
        ' passthrough — use loadfromtext which preserves all query properties
        oapp.loadfromtext acquery, queryname, qryfilepath
    else
        ' standard query — createquerydef from raw sql is much faster
        dim filenum as integer, sqltext as string
        filenum = freefile()
        open sqlfilepath for input as #filenum
        sqltext = input$(lof(filenum), filenum)
        close #filenum
        
        on error resume next
        oapp.currentdb.querydefs.delete queryname
        on error goto 0
        
        oapp.currentdb.createquerydef queryname, sqltext
    end if
    
    set fso = nothing
end sub

function splitformfile(filelocation)

dim dataline as string
dim codeline as boolean
codeline = false

dim myfile as string
myfile = replace(filelocation, ".form", ".bas")

dim fnout as integer, fnin as integer
fnout = freefile()
open myfile for output as #fnout

fnin = freefile()
open filelocation for input as #fnin

while not eof(fnin)
    line input #fnin, dataline
    
    if codeline then
        print #fnout, dataline
    end if
    
    if dataline = "CodeBehindForm" then codeline = true
wend

close #fnin
close #fnout

end function

sub normalizeexportfile(filepath as string)
    ' normalize casing in saveastext output to prevent false diffs.
    ' access randomly changes identifier casing between exports.
    ' this lowercases everything except content inside double quotes,
    ' preserving user-visible strings like msgbox text.
    
    dim fnin as integer, fnout as integer
    dim tmppath as string, dataline as string
    
    tmppath = filepath & ".tmp"
    
    fnin = freefile()
    open filepath for input as #fnin
    fnout = freefile()
    open tmppath for output as #fnout
    
    while not eof(fnin)
        line input #fnin, dataline
        dataline = lcasepreservestrings(dataline)
        print #fnout, dataline
    wend
    
    close #fnin
    close #fnout
    
    kill filepath
    name tmppath as filepath
    
end sub

function lcasepreservestrings(byval s as string) as string
    ' lowercases everything outside of double-quoted strings.
    ' "Hello World" stays as-is, but dim myvar becomes dim myvar.
    dim result as string, i as long, inquote as boolean, ch as string
    
    result = ""
    inquote = false
    
    for i = 1 to len(s)
        ch = mid$(s, i, 1)
        if ch = """" then
            inquote = not inquote
            result = result & ch
        elseif inquote then
            result = result & ch  ' preserve original case inside quotes
        else
            result = result & lcase$(ch)
        end if
    next i
    
    lcasepreservestrings = result
end function

function addnote(notetxt as string, optional batch as boolean = false)

currentdb.execute "INSERT INTO tblReleaseTracking(task) VALUES('" & strquotereplace(notetxt) & " ')", dbfailonerror

' in batch mode, skip the expensive requery/scroll — caller will flush at the end
if batch then exit function

on error resume next
form_sfrmtracking.requery
movetrackingtolastrecord

doevents

end function

sub flushnotes()
' call after a batch of addnote calls to update the tracking display once
    on error resume next
    form_sfrmtracking.requery
    movetrackingtolastrecord
    doevents
end sub
