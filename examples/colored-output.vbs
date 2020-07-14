option explicit


dim wsh
set wsh = createObject("wScript.shell")

dim use_ps_trick

'
'   Try to read registry value. If value does not exist, an error is thrown.
'   Therefore, we use «on error resume next» to catch such an error.
'
dim virtualTerminalLevel
on error resume next
virtualTerminalLevel = wsh.regRead("HKCU\Console\VirtualTerminalLevel")

if err.number <> 0 then ' {
 '
 ' The registry value does not exist, we have to
 ' use the «powershell trick»
 '
   use_ps_trick = true
else
 '
 ' The registry valued does exist. We use
 ' the «powershell trick» depending on the value:
 '
   if virtualTerminalLevel = 0 then
      use_ps_trick = true
   else
      use_ps_trick = false
   end if

end if ' }

'
'  Go back to normal error handling:
'
on error goto 0

if use_ps_trick then ' {

   dim ps
   set ps = wsh.exec("powershell.exe -noProfile -executionPolicy bypass -c ""exit""")
   while ps.status = 0
         wScript.sleep 50
   wend

end if ' }

dim txt
txt = chr(27) & "[91mRed" & chr(27) & "[0m normal"

wscript.echo txt

dim fso, stdOut
set fso = createObject("scripting.fileSystemObject")
set stdOut = fso.getStandardStream(1)

stdOut.writeLine(txt)
stdOut.writeLine("[91mRed[0m normal")


'
'    Overwriting existing text:
'
'       <esc>[1F  is: move cursor up one line
'
'    Thus, together with wscript.sleep
'    existing text is overwritten each second
'    with a new color.
'

wscript.echo ""
wscript.echo      "[92mGreen[0m  "

wscript.sleep 1000
wscript.echo "[1F[93mYellow[0m "

wscript.sleep 1000
wscript.echo "[1F[94mBlue[0m   "

wscript.sleep 1000
wscript.echo "[1F[95mMagenta[0m"
