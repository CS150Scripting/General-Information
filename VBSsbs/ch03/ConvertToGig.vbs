'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/14/2006
'
' NAME: convertToGig.vbs
'
' COMMENT: Key concepts are listed below:
'1.Uses If ... Then statement to make decision
'2.Converts to Gigabytes
'3.Converts to MegaBytes
'4.Uses wscript.quit to exit the if ... then statement.
'==========================================================================
Option Explicit
Dim intMemory
intMemory = 120000
intMemory = formatNumber(intMemory/1024)
	If intMemory > 1024 Then
		intMemory = formatNumber(intMemory/1024) & " Gigabytes"
			WScript.Echo intMemory
			WScript.quit
	End If
	intMemory = intMemory & " Megabytes"
WScript.Echo intMemory