'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/5/2006
'
' NAME: DisplayComputerNames.vbs
'
' COMMENT: Key concepts are listed below:
'1.Uses the WshSHell object to get access to regRead method. The WshShell Is
'2.Created when we use createObject("WScript.Shell")
'3.Note the use of line concatenation and continuation here. & _
'==========================================================================
Option Explicit     
On Error Resume Next
					
Dim objShell
Dim regActiveComputerName, regComputerName, regHostname
Dim ActiveComputerName, ComputerName, Hostname


regActiveComputerName = "HKLM\SYSTEM\CurrentControlSet" &_ 
	"\Control\ComputerName\ActiveComputerName\ComputerName"
regComputerName = "HKLM\SYSTEM\CurrentControlSet\Control" &_
	"\ComputerName\ComputerName\ComputerName"
regHostname = "HKLM\SYSTEM\CurrentControlSet\Services" &_
	"\Tcpip\Parameters\Hostname"

Set objShell = CreateObject("WScript.Shell")
ActiveComputerName = objShell.RegRead(regActiveComputerName)
ComputerName = objShell.RegRead(regComputerName)
Hostname = objShell.RegRead(regHostname)

WScript.Echo activecomputername & " is active computer name"
WScript.Echo ComputerName & " is computer name"
WScript.Echo Hostname & " is host name"
