'this uses the new win32_PingStatus wmi class to fake the standard ping
'command. It follows RFC 791. Check out the Feb. 2003 platform SDK 
' WMI chapter for all the details. 

' we set initial range of the network to search
' then we use For i = 1 to 255 
' then we concatenate the first three octets, with the number from for i

strMachines = "127.0.0."
For i = 1 To 25 Step 5
aMachines = strMachines & i

Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
ExecQuery("select * from Win32_PingStatus where address = '"_
& amachines & "'")
For Each objStatus in objPing
	If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
	WScript.Echo("machine " & amachines & " is not reachable") 
	else
	wscript.Echo("reply from " & amachines) 
	End If
Next
Next



