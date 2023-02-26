dim i,j
const one_sec = 1000
Set Shell = CreateObject("WScript.Shell")

For j = 1 To 2
	For i = 1 to 3
	Shell.Run "%comspec% /c echo " & Chr(07), 0, True
	Next
wscript.sleep one_sec
Next