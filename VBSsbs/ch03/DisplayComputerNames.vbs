Option Explicit     
'On Error Resume Next
					
Dim objShell
Dim regActiveComputerName, regComputerName, regHostname
Dim ActiveComputerName, ComputerName, Hostname


regActiveComputerName = "HKLM\SYSTEM\CurrentControlSet" &_ 
	"\Control\ComputerName\ActiveComputerName\ComputerName"
regComputerName = "HKLM\SYSTEM\CurrentControlSet\Control" &_
	"\ComputerName\ComputerName\ComputerName"
regHostname = "HKLM\SYSTEM\CurrentControlSet\Services" &_
	"\Tcpip\Parameters\Hostname"

Set objShell = WScript.CreateObject("WScript.Shell")
ActiveComputerName = objShell.RegRead(regActiveComputerName)
ComputerName = objShell.RegRead(regComputerName)
Hostname = objShell.RegRead(regHostname)

WScript.Echo activecomputername & " is active computer name"
WScript.Echo ComputerName & " is computer name"
WScript.Echo Hostname & " is host name"
