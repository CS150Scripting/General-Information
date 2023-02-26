'==========================================================================
'
' AUTHOR: Ed Wilson , MS,  10/27/2005
'
' NAME: WorkWithAsciiValues.vbs
'
' COMMENT: Key concepts are listed below:
'1.You can obtain the hex value using charmap. it shows up in the lower corner
'2.When we do a binary compare, we are comparing the binary representation of a 
'3. ascii value of the letters. This is, of course, case sensitive.
'==========================================================================


WScript.Echo "The letter assocated with ASCII 83 "_
	& "using chr() is: "  & Chr(83)
WScript.Echo "The ascii value of S " _
	& " using ASC() is: " & Asc("S")
WScript.Echo "The hexidecimal value of S " _
	& "using Hex(asc()) is: " & Hex(Asc("S"))
