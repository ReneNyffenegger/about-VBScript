option explicit

const  foo       =   42
const  bar       = &H80
const  baz_date  = #08/28/1970#
const  baz_time  = #22:23:24#

dim    baz
baz = baz_date + baz_time

wscript.echo "foo =      " & foo
wscript.echo "bar =      " & bar
wscript.echo "baz_date = " & baz_date
wscript.echo "baz_time = " & baz_time
wscript.echo "baz =      " & baz
