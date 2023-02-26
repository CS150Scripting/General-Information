'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/12/2006
'
' NAME: <MonitorForChangedDiskSpace.vbs>
'
' COMMENT: Key concepts are listed below:
'1.Uses an instanceModification event driven query
'2.Uses a do while loop to provide the repetitive checking
'3.Specifies to check for an event within 10 seconds.
'4.Uses timer function to monitor running time of the script.
'5.Uses formatNumber function to clean up the number returned by timer.
'==========================================================================
Option Explicit 
'On Error Resume Next 
Dim colMonitoredDisks
Dim objWMIService
Dim objDiskChange
Dim strComputer
Dim startTime, snapTime   'Used for timer Function

Const LOCAL_HARD_DISK = 3 'The driveType value from SDK
Const RUN_TIME = 10 	  'Time to allow the script to run in seconds
strComputer = "."
startTime = Timer

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colMonitoredDisks = objWMIService.ExecNotificationQuery _
     ("Select * from __instancemodificationevent within 10 where " _
         & "TargetInstance ISA 'Win32_LogicalDisk'")
Do While True 
snapTime = Timer
  Set objDiskChange = colMonitoredDisks.NextEvent
    If objDiskChange.TargetInstance.DriveType = LOCAL_HARD_DISK Then
   		WScript.echo "diskSpace on " &_
   			objDiskChange.TargetInstance.deviceID &_
   				" has changed. It now has " &_
   			objDiskChange.TargetInstance.freespace &_
   				" Bytes free."
    End If
    	If (snapTime - startTime) > RUN_TIME Then
   			Exit Do
   		End If 
Loop
WScript.Echo FormatNumber(snapTime-startTime) & " seconds elasped. Exiting now"
WScript.quit