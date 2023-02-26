'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <ParseAPPlog-Commented.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 9/17/2003
'
' COMMENT: <This script demonstrates using two InStr commands to populate an array. 
' It then uses the split command to create a multi-dimensional array that is used
' to customize the message obtained from the appLog file.>
'
'==========================================================================
' header information section
Option Explicit
On Error Resume Next
Dim arrTxtArray()'Dynamic array
Dim appLog 	 'Holds name of the application log
Dim SearchString 'First search string
Dim objTextFile 
Dim strNextLine 
Dim intSize 
Dim objFSO 
Dim i 
Dim ErrorString 'Used for second search string
Dim newArray 'New array created to sort output
' reference information section
intSize = 0 
appLog = "applog.csv" 	'Ensure in path
SearchString = "," 	'We search for each line containing a ,
ErrorString = "1004" 	'This particular error is from MSI Installer. Easily changable here.
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (appLog, ForReading)
' worker information section
Do until objTextFile.AtEndOfStream 'Using do until here is easier than do while <> 
    strNextLine = objTextFile.Readline
    if InStr (strNextLine, SearchString)Then 'If we find a , then we go no next line
    	If InStr ( strNextLine, ErrorString) Then 'Now we filter out our error message
	    	ReDim Preserve arrTxtArray(intSize) 'Now we resize the array
	    	arrTxtArray(intSize) = strNextLine
	    	intSize = intSize +1
	    End if
    End if
Loop
	objTextFile.close
' output information section 
For i = LBound(arrTxtArray) To UBound(arrTxtArray) 'This avoids errors by specifying the upper and lower boundaries
	If InStr (arrTxtArray(i), ",") Then 	   'Once again we look for commas
	newArray = Split (arrTxtArray(i), ",") 	   'Now we are going to split the array elements from first array
		WScript.Echo "Date: " & newArray(0) 'Each field from the application log is seperated by a comma
		WScript.Echo "Time: " & newArray(1) 'We simply choose the fields and the arrangement we want.
		WScript.Echo "Source: " & newArray(2)& " "& newArray(3) ' and we leave out the fields we dont want
		WScript.Echo "Server: " & newArray(7) 
		WScript.Echo "Message1: " & newArray(8)'Could further cleanup the message if you wish. 
		WScript.Echo "Message2: " & newArray(9)
		WScript.Echo "Message3: " & newArray(10)
		WScript.Echo " "
	End if
Next
WScript.Echo("all done")