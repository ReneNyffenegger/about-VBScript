option explicit

dim wsh
set wsh = createObject("wScript.shell")

dim userEnv
dim systemEnv
dim processEnv
dim volatileEnv

set userEnv     = wsh.environment("user"    )
set systemEnv   = wsh.environment("system"  )
set processEnv  = wsh.environment("process" )
set volatileEnv = wsh.environment("volatile")

wscript.echo("user's temp  = " & userEnv   ("temp"        ))
wscript.echo("systemRoot   = " & systemEnv ("os"          ))
wscript.echo("computername = " & processEnv("computername"))

'
'  Expand strings that contain %VARNAME%:
'
wScript.echo(wsh.expandEnvironmentStrings("temp         = %temp%"))

wscript.echo()
wscript.echo("Volatile variables:")
dim varName
for each varName in volatileEnv
    wScript.echo("  " & varName)
next
