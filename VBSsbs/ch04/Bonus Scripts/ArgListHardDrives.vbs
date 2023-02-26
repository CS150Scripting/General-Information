'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  5/12/2003
'
' NAME: <ListHardDrives.vbs>
'
' COMMENT: Key concepts are listed below:
'1. Main purpose is to show how to use the for-next loop
'2. uses Constants for disk drive
'3. Shows how to connect to WMI, and execute a query
'4. Shows how to use a constant in a WMI query
'5. Demonstrates line continuation technique
'==========================================================================

Option Explicit    	 ' is used to force the scripter to declare variables
'On Error Resume Next ' is used to tell vbscript to go to the next line If
' it encounters an error
' Dim is used to declare varable names that are used in the script
Const DriveType = 3 ' used by WMI for fixed disks
' other drive types are 2 for removable, 4 for Network, 5 for CD
Dim colDrives 'Holder for what comes back from the WMI query
Dim drive 'Holder for name of each logical drive in colDrives 
Dim strMachines
Dim aMachines
Dim machine

If WScript.Arguments.count = 0 Then
WScript.Echo("You must enter a computer name")
else
strMachines = WScript.Arguments.Item(0)
aMachines = split(strMachines, ",")
	For Each machine in aMachines
	set coldrives =_ 
	GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
		("select DeviceID from Win32_LogicalDisk where DriveType =" & DriveType)
			for each drive in colDrives
			WScript.Echo drive.DeviceID
		Next
	Next
End if