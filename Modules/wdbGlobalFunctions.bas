option compare database
option explicit

declare ptrsafe function shellexecute lib "shell32.dll" alias "ShellExecuteA" (byval hwnd as long, byval lpoperation as string, byval lpfile as string, byval lpparameters as string, byval lpdirectory as string, byval lpnshowcmd as long) as long

function tojson(byval dict as object) as string
    dim key as variant, result as string, value as string

    result = "{"
    for each key in dict.keys
        result = result & iif(len(result) > 1, ",", "")

        if typename(dict(key)) = "Dictionary" then
            value = tojson(dict(key))
            tojson = value
        else
            value = """" & dict(key) & """"
        end if

        result = result & """" & key & """:" & value & ""
    next key
    result = result & "}"

    tojson = result
end function



public function genemail(optional byval strto as string = "", optional byval strbcc as string = "", optional byval strcc as string = "", optional byval strsubject as string = "", optional body as string = "") as boolean
genemail = false
    
dim objemail as object

set objemail = createobject("outlook.Application")
set objemail = objemail.createitem(0)

with objemail
    .to = strto
    .cc = strcc
    .bcc = strbcc
    .subject = strsubject
    .htmlbody = body
    .display
end with

set objemail = nothing

genemail = true
end function

public sub openpath(path)
createobject("Shell.Application").open cvar(path)
end sub

function emailcontentgen(subject as string, title as string, subtitle as string, primarymessage as string, detail1 as string, detail2 as string, detail3 as string) as string
emailcontentgen = subject & "," & title & "," & subtitle & "," & primarymessage & "," & detail1 & "," & detail2 & "," & detail3
end function

function getemail(username as string) as string
on error resume next

dim db as database
set db = currentdb()

dim rspermissions as recordset
set rspermissions = db.openrecordset("SELECT * from tblDeveloperInfo WHERE user = '" & username & "'")
getemail = rspermissions!email
rspermissions.close

db.close

end function

function ap_disableshift()

on error goto errdisableshift
dim db as dao.database
dim prop as dao.property
const conpropnotfound = 3270

set db = currentdb()

db.properties("AllowByPassKey") = false
exit function

errdisableshift:
if err = conpropnotfound then
set prop = db.createproperty("AllowByPassKey", dbboolean, false)
db.properties.append prop
resume next
else
msgbox "Function 'ap_DisableShift' did not complete successfully."
exit function
end if

end function

public function strquotereplace(strvalue)

strquotereplace = replace(nz(strvalue, ""), "'", "''")

end function

function ap_enableshift()

on error goto errenableshift
dim db as dao.database
dim prop as dao.property
const conpropnotfound = 3270

set db = currentdb()
db.properties("AllowByPassKey") = true
exit function

errenableshift:
if err = conpropnotfound then
set prop = db.createproperty("AllowByPassKey", dbboolean, true)
db.properties.append prop
resume next
else
msgbox "Function 'ap_DisableShift' did not complete successfully."
exit function
end if

end function
