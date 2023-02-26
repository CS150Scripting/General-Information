'==========================================================================
'
'
' NAME: <ListAllFolders.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 1/23/200
'
' COMMENT: <Uses WMI win32_Directory namespace to list all folders on computer.>
'  based on my WMI template
'==========================================================================

Option Explicit 
On Error Resume Next
dim strComputer
dim wmiNS
dim wmiQuery
dim objWMIService
dim colItems
dim objItem

strComputer = "."
wmiNS = "\root\cimv2"
wmiQuery = "Select name from win32_directory"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecQuery(wmiQuery)

For Each objItem in colItems
    Wscript.Echo ": " & objItem.name
    
Next