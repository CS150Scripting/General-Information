'==========================================================================
'
' NAME: <SearchTXT-Commented.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 9/16/2003
'
' <COMMENT: This script creates a dynamic array to hold lines parsed from
'  	a setup log file. Scripting points demonstrated are listed below:>
'1.Creating a dynamic Array
'2.Setting the initial size of an Array
'3.Connecting to the file system object to read a txt file
'4.Using InStr to parse a txt stream for a specific word
'5.Using the Do Until loop to iterate through a txt file
'6.Using the reDim Preserve to change the size of a dynamic Array
'7.Using LBound and UBound to iterate through an array
'==========================================================================
' The section below is the Header section of the script
Option Explicit
On Error Resume Next
Dim arrTxtArray()'Declares a dynamic array
Dim myFile 	 'Holds the file to open up
Dim SearchString 'Holds the string to search for
Dim objTextFile  'Holds the connection to the text file
Dim strNextLine  'Holds next line in the text stream
Dim intSize 	 'Holds the initial size of the array
Dim objFSO 	 'Holds connection to the file system object
Dim i 		 'Used to increment intSize counter

' The section below is the Reference section of the script
intSize = 0 'Used for initial size of the array
myFile = "c:\windows\setuplog.txt" 'Modify as required
SearchString = "Error"
Const ForReading = 1 'Tells filesystemObject we will read the file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (myFile, ForReading)
    
' The section below is the worker section of the script  
Do until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline
    if InStr (strNextLine, SearchString)then
    	ReDim Preserve arrTxtArray(intSize)
    	arrTxtArray(intSize) = strNextLine
    	intSize = intSize +1
    End if
Loop

objTextFile.close

'The section below is the the output section of the script
For i = LBound(arrTxtArray) To UBound(arrTxtArray)
	WScript.Echo arrTxtArray(i)
Next

WScript.Echo("all done")


