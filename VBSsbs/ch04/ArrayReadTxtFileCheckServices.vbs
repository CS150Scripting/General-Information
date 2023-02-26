'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/26/2006
'
' NAME: <ArrayReadTxtFileCheckServices.vbs>
' ver.1.2
' COMMENT: Key concepts are listed below:
'1.Creates an array of servers and services from a text file
'2.Uses WMI to check on the status of those services on those servers
'3.Uses Win32_Service wmi class.
'4. uses file system object to open up a text file and read it. 
'5. The txt file MUST not have any spaces between the commas, or script will fail
'==========================================================================
Option Explicit
'On Error Resume Next 
Dim objFSO ' holds connection to file system object
Dim objTextFile ' holds hook to the text file
Dim arrServiceList ' produced by split command. an array of servers and services
Dim strNextLine ' produced by reading a line from the text file
Dim i ' counter variable for looping operations. nothing exciting
Dim TxtFile ' holds path to text file that has names of servers and services
dim boundary ' upper boundary of array. Changes with each line of text file
Dim strComputer ' comes from element (0) in on each line of text file
Dim ServiceName ' comes from element (1) to upperboundary (ubound)
Dim wmiRoot ' target of WMI operation, includes moniker. 
Dim wmiQuery ' carefully paramaterized query. Variable from array for the name of service
Dim objWMIService ' hook into the WMI repository
Dim objItem ' the collection of services that comes back from wmi query.

TxtFile = "RealServersAndServices.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(TxtFile)
 
Do until  objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrServiceList = Split(strNextLine , ",")
    boundary = Ubound(arrServiceList)
    strComputer = arrServiceList(0)
    WScript.echo "Status of services on " & strComputer 	
    	SubcheckWMI
    WScript.echo vbNewLine
Loop

WScript.Echo("all done")

' *** subs and functions are below *****
Sub SubcheckWMI
wmiRoot = "winmgmts:\\" & strComputer & "\root\cimv2" 
Set objWMIService = GetObject(wmiRoot)
	For i = 1 to boundary
		ServiceName = arrServiceList(i)
		wmiQuery = "win32_service.name=" & funFIX(serviceName)
    Set objItem = objWMIService.get (wmiQuery)
			WScript.Echo vbtab & (servicename) & " Is: " _
			& objItem.state & " start up type is: " & objItem.StartMode
	Next
End Sub

Function funFIX(strIN)
	funFIX = "'" & strIN & "'"
End Function