option compare database
option explicit

function getbestsqldriver() as string

    dim shell as object
    dim driverlist as variant, driver as variant
    dim regpath as string
    
    set shell = createobject("WScript.Shell")
    ' list drivers in order of preference (newest first)
    driverlist = array("ODBC Driver 18 for SQL Server", _
                       "ODBC Driver 17 for SQL Server", _
                       "ODBC Driver 13 for SQL Server", _
                       "SQL Server Native Client 11.0", _
                       "SQL Server")
    
    on error resume next
    for each driver in driverlist
        ' check registry for the driver entry
        if regkeyexists("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & driver & "\Driver") then
            getbestsqldriver = driver
            exit function
        end if
    next driver
    
    getbestsqldriver = "" ' no driver found
    
end function

function regkeyexists(regpath as string) as boolean
    dim shell as object
    set shell = createobject("WScript.Shell")
    on error resume next
    shell.regread regpath
    regkeyexists = (err.number = 0)
    on error goto 0
end function

function relinksqltables(optional returnstringonly as boolean = false) as string

    dim db as dao.database
    dim tdf as dao.tabledef
    dim qdf as dao.querydef
    dim strdriver as string
    dim strconn as string
    
    strdriver = getbestsqldriver()
    if strdriver = "" then exit function
    
    set db = currentdb
    ' base odbc connection string
    
    dim strextra as string
    select case lcase$(strdriver)
        case lcase$("ODBC Driver 18 for SQL Server")
            strextra = ";Encrypt=Yes;TrustServerCertificate=Yes"
            
        case lcase$("ODBC Driver 17 for SQL Server")
            strextra = ";Encrypt=Yes;TrustServerCertificate=Yes"
            
        case else
            strextra = ""
    end select
    
    strconn = "ODBC;" & _
              "DRIVER=" & strdriver & ";" & _
              "SERVER=ITI-SQL\ITISQL;" & _
              "DATABASE=workingdb;" & _
              "Trusted_Connection=Yes;" & _
              "APP=Microsoft Office" & _
              strextra & ";"
    
    if returnstringonly then
        relinksqltables = strconn
        exit function
    end if
    
    ' loop through all tables and update odbc links
    for each tdf in db.tabledefs
        ' only relink tables that already have an odbc connection string
        if instr(1, tdf.connect, "SERVER=ITI-SQL") then
            tdf.connect = strconn
            tdf.refreshlink
        end if
    next tdf
    
    ' relink pass-through queries
    for each qdf in db.querydefs
        if qdf.type = dbqsqlpassthrough then
            if instr(1, qdf.connect, "SERVER=ITI-SQL") > 0 then
                qdf.connect = strconn
            end if
        end if
    next qdf
    
relinksqltables = strconn
    
end function
