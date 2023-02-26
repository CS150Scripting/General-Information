' script two in vbscript book
' Key concepts are listed below:
' 1. Option Explicit
' 2. On Error Resume Next
' 3. Declaring variable names
' 4. Assigning variable names to be equal to something
' 5. Creating the shell
' 6. Reading the registry
' 7. Making pop-up boxes

' This script displays User Information by reading the registry

Option Explicit     'forces the scripter to declare variables
On Error Resume Next 'tells vbscript to go to the next line 
					'instead of exiting when an error occurs

' Dim is used to declare varable names that are used in the script
Dim objShell
Dim regLogonUserName, regExchangeDomain, regGPServer
Dim regLogonServer, regDNSdomain
Dim LogonUserName, ExchangeDomain, GPServer
Dim LogonServer, DNSdomain

' when you use  a varable name and then the equals sign (=) 
'you are saying the varable is the same as what you just set it to.
'Since the registry keys are quite long, we split the line in two.
'This is done by closing the quotes on the first line using the & _
'characters and opening the quote on the next line. 
regLogonUserName = "HKEY_CURRENT_USER\Software\Microsoft\" & _
	"Windows\CurrentVersion\Explorer\Logon User Name"
regExchangeDomain = "HKEY_CURRENT_USER\Software\Microsoft\" & _
	"Exchange\LogonDomain"
regGPServer = "HKEY_CURRENT_USER\Software\Microsoft\Windows\" & _
	"CurrentVersion\Group Policy\History\DCName"
regLogonServer = "HKEY_CURRENT_USER\Volatile Environment\" & _
	"LOGONSERVER"
regDNSdomain = "HKEY_CURRENT_USER\Volatile Environment\" & _
	"USERDNSDOMAIN"

Set objShell = WScript.CreateObject("WScript.Shell")


LogonUserName = objShell.RegRead(regLogonUserName)
ExchangeDomain= objShell.RegRead(regExchangeDomain)
GPServer = objShell.RegRead(regGPServer)
LogonServer = objShell.RegRead(regLogonServer)
DNSdomain = objShell.RegRead(regDNSdomain)

' in order to make Dialog boxes you can use wscript.echo
' and then tell it what you want it to say. 
WScript.Echo LogonUserName & " is currently Logged on"
WScript.Echo ExchangeDomain & " is the current logon domain"
WScript.Echo GPServer & " is the current Group Policy Server"
WScript.Echo LogonServer & " is the current logon server"
WScript.Echo DNSdomain & " is the current DNS domain"
