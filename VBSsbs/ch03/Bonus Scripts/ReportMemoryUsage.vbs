'==========================================================================
'
' NAME: <ReportMemoryUsage.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/14/2006
'
' COMMENT: <Use the win32_operatingSystem class>
'1. Uses execQuery to return a wmi collection
'2. Uses for each next to walk throught the collection
'3. Uses a user defined function to translate from kilobytes to meg and gig
'4. Uses if then to make decision if meg or gig.
'==========================================================================

Option Explicit 
'On Error Resume Next
dim strComputer
dim wmiNS
dim wmiQuery
dim objWMIService
dim colItems
dim objItem

strComputer = "."
wmiNS = "\root\cimv2"
wmiQuery = "Select * from win32_OperatingSystem"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecQuery(wmiQuery)

For Each objItem in colItems
    WScript.Echo "FreePhysicalMemory: " & funConvert(objItem.FreePhysicalMemory)
    WScript.Echo "FreeSpaceInPagingFiles: " & funConvert(objItem.FreeSpaceInPagingFiles) 
    WScript.Echo "FreeVirtualMemory: " & funConvert(objItem.FreeVirtualMemory)
    WScript.Echo "MaxProcessMemorySize: " & funConvert(objItem.MaxProcessMemorySize) 
    WScript.Echo "TotalVirtualMemorySize: " & funConvert(objItem.TotalVirtualMemorySize) 
    WScript.Echo "TotalVisibleMemorySize: " & funConvert(objItem.TotalVisibleMemorySize) 
Next

Function funConvert(intFunction)
Dim intMemory
intMemory = formatNumber(intFunction/1024)
	If intMemory > 1024 Then
		funConvert = formatNumber(intMemory/1024) & " Gigabytes"
	End If
	funConvert = intMemory & " Megabytes"
End Function