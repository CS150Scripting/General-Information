'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <SelectCase.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2003
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
Dim colComputers
Dim objComputer
Dim strComputerRole
Dim strcomputerName
Dim strDomainName
Dim strUserName
strComputer = "."
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
wmiQuery = "Select * from win32_computersystem"
Set objWMIService = GetObject(wmiRoot)
Set colComputers = objWMIService.ExecQuery _
    (wmiQuery)
For Each objComputer in colComputers
strComputerName = objComputer.name
strDomainName = objComputer.Domain
strUserName = objComputer.UserName
    Select Case objComputer.DomainRole 
        Case 0 
            strComputerRole = "Standalone Workstation"
        Case 1        
            strComputerRole = "Member Workstation"
        Case 2
            strComputerRole = "Standalone Server"
        Case 3
            strComputerRole = "Member Server"
        Case 4
            strComputerRole = "Backup Domain Controller"
        Case 5
            strComputerRole = "Primary Domain Controller"
    End Select
    WScript.Echo strComputerRole & vbcrlf & strComputerName & vbcrlf & strDomainName & vbcrlf & strUserName
Next
WScript.Echo("all done")

