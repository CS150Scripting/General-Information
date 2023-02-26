'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  2/18/2006
' version: 2.0
' NAME: <CustomizeExistingScriptSolution.vbs>
'
' COMMENT: Key concepts are listed below:
'1. reading from the registry: crash recovery information
'2. Declaring variables
'3. Option Explicit
'4. On Error Resume Next
'5. wscript.shell, and wscript.CreateObject
'6. Set command
'7. wscript.echo command. 
'==========================================================================
Option Explicit ' forces declaring of variables    
On Error Resume Next ' tells vbscript to go to the next line in the code
					
Dim objShell 'holds hook for access to wscript.shell object
Dim regAutoBoot, regMiniDump, regHostname, regLogEvent,regDumpFile  ' holds registry keys
Dim AutoBoot, MiniDump, Hostname, LogEvent,DumpFile ' used to echo out the registry values

regAutoBoot = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\AutoReboot"
regMiniDump = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\MinidumpDir"
regHostname = "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Hostname"
regLogEvent = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\LogEvent"
regDumpFile = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\DumpFile"


Set objShell = wscript.CreateObject("WScript.Shell") ' we use set, and createObject to allow access to the shell

AutoBoot = objShell.RegRead(regAutoBoot) ' to read from registry we use regRead. To write, we use regWrite. 
MiniDump = objShell.RegRead(regMiniDump)
Hostname = objShell.RegRead(regHostname)
LogEvent = objShell.RegRead(regLogEvent)
DumpFile = objShell.RegRead(regDumpFile)

WScript.Echo AutoBoot & " is autoboot configuration" ' the & is called the concatenation symbol
WScript.Echo MiniDump & " is miniDump config"        ' which means we are smushing stuff together. 
WScript.Echo Hostname & " is host name"
WScript.Echo LogEvent & " is logEvent configuration"
WScript.Echo DumpFile & " is dumpFile"