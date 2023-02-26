'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <enableDHCPSolution.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/26/2003
'
' COMMENT: <The following items are taught:>
' 1. WMI queries
' 2. For Each Next
' 3. If Then Else
' 4. Use of Variables
'
'==========================================================================

Option Explicit
On Error Resume Next 
Dim strComputer
Dim wmiRoot
Dim wmiQuery
Dim objWMIService
Dim colNetAdapters
Dim objNetAdapter
Dim DHCPEnabled
strComputer = "."
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2"
wmiQuery = "Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE"
Set objWMIService = GetObject(wmiRoot)
Set colNetAdapters = objWMIService.ExecQuery _
    (wmiQuery)
For Each objNetAdapter In colNetAdapters
    DHCPEnabled = objNetAdapter.EnableDHCP()
     If DHCPEnabled = 0 Then
        WScript.Echo "DHCP has been enabled."
    Else
        WScript.Echo "DHCP could not be enabled."
    End If
Next
