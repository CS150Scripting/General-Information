'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ParseDayLightSavingsTimeArray.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 5/19/2004
'
' COMMENT: <Following Concepts Presented>
' 1. Declaring an Array
' 2. adding information to an array
' 3. tracking the progress of script execution
'==========================================================================
Option Explicit 
Dim arDOW, arMonth
Dim strComputer
Dim objWmiService
Dim wmiNS
Dim wmiQuery
Dim objItem
Dim colItems

arDOW = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
arMonth = Array("January", "Feburary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
strComputer = "."
wmiNS = "\root\cimv2"
wmiQuery = "Select * from Win32_TimeZone"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS)
Set colItems = objWMIService.ExecQuery(wmiQuery)

For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    WScript.echo "Day of Week setting is: " & objItem.dayLightDayOfWeek & " which is: " & arDOW(objItem.DaylightDayOfWeek)
    WScript.echo "Hour: " & objItem.DaylightHour 
    WScript.echo "Month: " & objItem.DaylightMonth & " which is: " & arMonth(objItem.DaylightMonth -1)
    WScript.echo "Description: " & objItem.DaylightName 
    WScript.echo "the transition from DLS to Standard occurs: " 
    WScript.echo "Day of Week setting is: " & objItem.standardDayOfWeek & " which is: " & arDOW(objItem.DaylightDayOfWeek)
    WScript.echo "Hour: " & objItem.StandardHour 
    WScript.echo "Month: " & objItem.StandardMonth & " which is: " & arMonth(objItem.StandardMonth -1)
    WScript.echo "Description: " & objItem.StandardName 
Next

