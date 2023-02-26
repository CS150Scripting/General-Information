'ed wilson 
'script three in vbscript book
' Key concepts are listed below:
' 1. Option Explicit
' 2. On Error Resume Next
' 3. Declaring variable names
' 4. Assigning variable names to be equal to something
' 5. Creating the shell
' 6. Creating a file system object
' 7. Reading the registry
' 8. Making pop-up boxes

'This script displays various Computer Names by reading the registry

Option Explicit     	'Forces the scripter to declare variables
On Error Resume Next 	'Tells vbscript to go to the next line 
			'instead of exiting when an error occurs

' Dim is used to declare varable names that are used in the script
Dim objShell
Dim regActiveComputerName, regComputerName, regHostname
Dim ActiveComputerName, ComputerName, Hostname

'When you use  a varable name and then the equals sign (=) 
'you are saying the varable is the same as what you just set it to.
'Since the registry keys are quite long, we split the line in two.
'This is done by closing the quotes on the first line using the &_
'characters and opening the quote on the next line. 

regActiveComputerName = "HKLM\SYSTEM\CurrentControlSet" &_ 
	"\Control\ComputerName\ActiveComputerName\ComputerName"
regComputerName = "HKLM\SYSTEM\CurrentControlSet\Control" &_
	"\ComputerName\ComputerName\ComputerName"
regHostname = "HKLM\SYSTEM\CurrentControlSet\Services" &_
	"\Tcpip\Parameters\Hostname"

'To read the registry, we create a shell object. Then we use the RegRead method
'create a wscript shell and set it equal to the variable objshell

Set objShell = WScript.CreateObject("WScript.Shell")
ActiveComputerName = objShell.RegRead(regActiveComputerName)
ComputerName = objShell.RegRead(regComputerName)
Hostname = objShell.RegRead(regHostname)

'In order to make pop-up windows boxes you can use wscript.echo
'and then tell it what you want it to say. 
WScript.Echo activecomputername & " is active computer name"
WScript.Echo ComputerName & " is computer name"
WScript.Echo Hostname & " is host name"