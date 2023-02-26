'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <workWith2DArray.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/8/2003
'
' COMMENT: <Following Concepts Presented>
' 1. Declaring a 2 dimmension Array
' 2. dynamically adding information to an array
' 3. tracking the progress of script execution
'==========================================================================
Option Explicit 
Dim i ' first element
Dim j ' second element
Dim numLoop ' counts loops 
Dim a (3,3) ' two dimension array with 4 elements each.
numLoop = 0
For i = 0 To 3
	For j = 0 To 3
numLoop = numLoop+1
WScript.Echo "i = " & i & " j = " & j
a(i, j) = "loop " & numLoop
WScript.Echo "Value stored in a(i,j) is: " & a(i,j)
	Next
Next
