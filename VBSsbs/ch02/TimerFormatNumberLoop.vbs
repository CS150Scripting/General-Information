'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/12/2006
'
' NAME: TimerFormatNumberLoop.vbs
'
' COMMENT: Key concepts are listed below:
'1.Use of the timer function
'2.Use of the formatNumber function
'3.Use of Do while loop 
'==========================================================================
option explicit
'On Error Resume Next
dim startTime,EndTime, TotalTime

const sleepTime = 1000
startTime = timer

Do While totalTime < 5
	wscript.echo "StartTime is: " & startTime
		endTime = timer
	wscript.echo "EndTime is: " & endtime
		TotalTime = EndTime - StartTime
	wscript.echo "TotalTime is: " & TotalTime
		totalTime = formatNumber(totaltime)
	wscript.sleep sleepTime
	wscript.echo vbnewline
Loop