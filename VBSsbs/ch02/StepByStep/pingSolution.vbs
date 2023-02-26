'this uses the new win32_PingStatus wmi class to fake the standard ping
'command. It follows RFC 791. Check out the Feb. 2003 platform SDK 
' WMI chapter for all the details. 

' header section
Option Explicit
On Error Resume next
Dim strMachines ' holds string of names of machines to ping
dim aMachines ' used to hold individual name of computer to ping
dim machine ' used to keep track of which computer is pinged
Dim i ' keeps track of how many pings are sent
Dim objPing ' connection to wmi to allow us to ping
Dim objStatus ' these are the status codes returned by ping

' reference section

strMachines = "127.0.0.1;localHost;127.0.0.2"
aMachines = split(strMachines, ";")

' worker and output section

For Each machine in aMachines
	For i = 1 To 3
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
		ExecQuery("select * from Win32_PingStatus where address = '"_
		& machine & "'")
			For Each objStatus in objPing
				If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
				WScript.Echo("machine " & machine & " is not reachable") 
				else
				wscript.Echo("reply from " & machine) 
				End If
			Next
	Next
Next


