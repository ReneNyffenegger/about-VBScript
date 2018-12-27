option explicit

wscript.echo "You gave " & wscript.arguments.count & " arguments."

dim argNo

for argNo = 1 to wScript.arguments.count
    wscript.echo "Argument " & argNo & " is: " & wScript.arguments(argNo-1)
next
