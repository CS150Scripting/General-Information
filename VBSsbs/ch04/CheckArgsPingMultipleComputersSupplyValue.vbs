'this uses the new win32_PingStatus wmi class to fake the standard ping
'command. It follows RFC 791. Check out the Feb. 2003 platform SDK 
' WMI chapter for all the details. 
'CheckArgsPingMultipleComputers.vbs

Option Explicit
Dim colargs, aMachines, machine
Dim objPing, objStatus
Dim strMachines

Set colargs = WScript.Arguments.UnNamed
subCheckArgs 'checks the count property for arguments


aMachines = split(strMachines, ";")
For Each machine in aMachines
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
ExecQuery("select * from Win32_PingStatus where address = '"_
& machine & "'")
For Each objStatus in objPing
	If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
	WScript.Echo("machine " & machine & " is not reachable") 
	Else
	wscript.Echo("reply from " & machine) 
	End If
Next
Next

Sub subCheckArgs
If colargs.count =0 Then
  WScript.Echo "You must enter a computer to ping" & VbCrLf & _
    "Try this: cscript CheckArgsPingMultipleComputers.vbs " _
	 & " 127.0.0.1;localhost" & VbCrLf & _
	 "Pinging default values ..." & vbcrlf 	
	strMachines = "127.0.0.1;localhost"
 Else
 strMachines = colargs(0)
End If
End Sub 


