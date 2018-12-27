option explicit

dim namedArgs
set namedArgs = wscript.arguments.named


dim paramName
for each paramName in array("foo", "bar", "baz")

  '
  ' Check if paramName was specified when script
  ' was invoked:
  '

    if namedArgs.exists(paramName) then
       wscript.echo "Value for " & paramName & " = " & namedArgs(paramName)
    else
       wscript.echo "No value given for " & paramName
    end if

next
