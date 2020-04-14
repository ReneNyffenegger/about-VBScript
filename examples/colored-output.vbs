option explicit

dim txt
txt = chr(27) & "[91mRed" & chr(27) & "[0m"

wscript.echo txt

dim fso, stdOut
set fso = createObject("scripting.fileSystemObject")
set stdOut = fso.getStandardStream(1)

stdOut.writeLine(txt)
stdOut.writeLine("[91mRed")
