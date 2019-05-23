option explicit

wscript.echo "1st non-optional paramter: " & wscript.arguments.unnamed(0)
wscript.echo "2nd non-optional paramter: " & wscript.arguments.unnamed(1)

if wscript.arguments.named.exists("opt") then
   wscript.echo "opt was chosen"
end if
