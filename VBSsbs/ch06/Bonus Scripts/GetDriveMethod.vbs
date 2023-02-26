'==========================================================================
'
' VBScript:  AUTHOR: Student , msft,  2/4/2004
'
' NAME: <GetDriveMethod.vbs>
'
' COMMENT: Key concepts are listed below:
'1.
'2.
'3. 
'==========================================================================
Option Explicit
Dim fso, c, s, driveLetter, showFreeSpace
driveLetter = "c:"

Set fso = CreateObject("Scripting.FileSystemObject")
Set c = fso.GetDrive(fso.GetDriveName(driveLetter))
s = "Drive " & UCase(driveLetter) & " - " 
s = s & c.VolumeName
s = s & "Free Space: " & FormatNumber(c.FreeSpace/1024, 0) 
s = s & " Kbytes"
ShowFreeSpace = s
WScript.Echo showfreespace
