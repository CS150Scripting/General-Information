'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ComputerRoles.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 3/16/2006
'Version 2.0 Added function. Added case Else. 
'
' COMMENT: <illustrates select case>
'
'==========================================================================
Option Explicit
'On Error Resume Next 
Dim strComputer
Dim wmiRoot
Dim wmiQuery
Dim objWMIService
Dim colItems
Dim objItem

strComputer = "."
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
wmiQuery = "Select DomainRole from Win32_ComputerSystem"

Set objWMIService = GetObject(wmiRoot)
Set colItems = objWMIService.ExecQuery _
    (wmiQuery)
For Each objItem in colItems
    WScript.Echo funComputerRole(objItem.DomainRole)
Next

Function funComputerRole(intIN)
  Select Case intIN
    Case 0 
        funComputerRole = "Standalone Workstation"
    Case 1        
        funComputerRole = "Member Workstation"
    Case 2
        funComputerRole = "Standalone Server"
    Case 3
        funComputerRole = "Member Server"
    Case 4
        funComputerRole = "Backup Domain Controller"
    Case 5
        funComputerRole = "Primary Domain Controller"
    Case Else
    	funComputerRole = "Look this one up in SDK"
  End Select
End Function