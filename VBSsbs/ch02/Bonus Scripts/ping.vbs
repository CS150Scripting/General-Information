'this uses the new win32_PingStatus wmi class to fake the standard ping
'command. It follows RFC 791. Check out the Feb. 2003 platform SDK 
' WMI chapter for all the details. 


strMachines = "s1;s2"
aMachines = split(strMachines, ";")
For Each machine in aMachines
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


