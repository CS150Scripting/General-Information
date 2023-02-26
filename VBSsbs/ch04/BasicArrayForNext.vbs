'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  6/25/2006
'
' NAME: <BasicArrayForNext.vbs>
'ver.1.2
' COMMENT: Key concepts are listed below:
'1.Create a static array
'2.Retrieve by element number
'3.Uses space function, ubound function, array function
'==========================================================================
Option Explicit    	    
On Error Resume Next 
Dim myTab 'Holds custom tab of two places
Dim aryComputer 'Holds array of computer names
Dim computer 	'Individual computer from the array
Dim i		'Simple counter variable. Used to retrieve by
		'Element number in the array. 
myTab = Space(2)
i = 0		'The first element in an array is 0.
aryComputer = array("s1","s2","s3")


WScript.Echo "Retrieve via for next"
i=0
For i = 0 To UBound(aryComputer)
	WScript.Echo myTab & "computer # " & i & _
		" is " & aryComputer(i)
Next

