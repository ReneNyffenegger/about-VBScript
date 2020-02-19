option explicit

dim t_start, t_end

t_start = timer

msgBox "Press OK when done"

t_end   = timer

wScript.echo "Elapsed time in seconds: " & formatNumber(t_end - t_start, 1)
