option compare database
option explicit

function setsplashloading(label as string)

if isnull(tempvars!loadamount) then exit function
tempvars.add "loadAmount", tempvars!loadamount + 1
form_frmsplash.lnloading.width = (tempvars!loadamount / 5) * tempvars!loadwd
form_frmsplash.lblloading.caption = label
form_frmsplash.repaint

end function

function assignthemetoparameters(themeid as long)

dim db as database
set db = currentdb

db.execute "UPDATE tblParameters SET themeId = " & themeid

set db = nothing

end function

function disableshift()

dim db, acc
set acc = createobject("Access.Application")
'set db = acc.dbengine.opendatabase("\\data\mdbdata\WorkingDB\build\Commands\Misc_Commands\WorkingDB_SummaryEmail.accdb", false, false)
'set db = acc.dbengine.opendatabase("H:\dev\WorkingDB_SummaryEmail.accdb", false, false)
set db = acc.dbengine.opendatabase("C:\workingdb\WorkingDB_ghost.accdb", false, false)


db.properties("AllowByPassKey") = true

db.close
set db = nothing

end function

function getpassword()

dim db as database
set db = opendatabase("")

dim rs as recordset
set rs = db.openrecordset("SELECT * FROM MSysObjects WHERE Connect is not null")

do while not rs.eof
    debug.print "Database: " & rs!database & vbtab & " Connection: " & rs!connect
    rs.movenext
loop

rs.close
set rs = nothing
db.close
set db = nothing

end function
