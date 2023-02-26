'==========================================================================
' NAME: <WhileWendLoop.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 3/12/2006
'
' COMMENT: <Use While ...Wend Loop>
'1. Uses While Wend Loop statement. 
'2. Uses the timeserial function to turn numbers into a timestamp
'3. Uses the time function to get current timestamp
'==========================================================================
Option Explicit 
'On Error Resume Next
dim dtmTime

Const hideWindow = 0
Const sleepyTime = 1000
dtmTime = timeSerial(19,25,00) 'Modify this value with desired time

while dtmTime > Time
	wscript.echo "current time is: " & Time &_
		" counting to " & dtmTime
	WScript.Sleep sleepyTime
Wend
subBeep
WScript.Echo dtmTime &  " was reached."


Sub subBeep
Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "%comspec% /c echo " & Chr(07),hideWindow
End Sub