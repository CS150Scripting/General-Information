'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/6/2006
'
' NAME: <ShowErrors.vbs>
'
' COMMENT: Key concepts are listed below:
'1.uses for next to loop through and raise a bunch of errors
'2.uses instr to filter out the errors that come back as unknown
'3.uses err.clear to clear the errors from the error object
'==========================================================================

Option Explicit    	'Is used to force the scripter to declare variables
On Error Resume Next 	'Is used to tell vbscript to go to the next line if it encounters an error

Dim i 'counter variable
For i = 1 To 500
Err.Raise i   
	If InStr(err.Description, "Unknown") = 0 Then
	WScript.echo ("Error # " & (Err.Number) & " " & Err.Description)
	End If 
	Err.Clear      'Clear the error.
Next
