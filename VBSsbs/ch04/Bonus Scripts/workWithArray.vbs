'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <workWithArray.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/8/2003
'
' COMMENT: <Following Concepts Presented>
' 1. Declaring a 4 element Array
' 2. dynamically adding information to an array
' 3. tracking the progress of script execution
'==========================================================================

Dim a (3)
For n = 0 To 3
WScript.Echo "n = " &(n)
a(i) = n
WScript.Echo "a(i) = " & a(i)
Next
