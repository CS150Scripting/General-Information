'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ifThenElse.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2003
'
' COMMENT: <shows how to use IF THEN ELSE>
'
'==========================================================================
Option Explicit
On Error Resume Next 
Dim a,b,c,d
a = 1
b = 2
c = 3 
d = 4
If a + b = d Then
WScript.Echo (a & " + " & b & " is equal to " & d)
Else
WScript.Echo (a & " + " & b & " is equal to " & c)
End If 