'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: <OneStepFurtherPT1.vbs>
'
' AUTHOR: ed wilson , mred
' DATE  : 6/26/2006
'
' COMMENT: <This script WILL NOT run, it is INCOMPLETE. It covers up to>
'					<Step 6 in the one step further exercise for chapter four>
'1. assigning wmi connection to wmiRoot
'2. assigning wmi query to wmiQuery variable
'3. For Each
'4. ReDim
'5. Use of Array FILTER command**
'6. Working with UBOUND and using for outputting info
'7. Line concatenation
'==========================================================================
Option Explicit
'On Error Resume next
Dim computer ' means this computer
Dim wmiRoot ' holds connection to wmi namespace
Dim objWMIService ' holds connection for wmi
Dim wmiQuery ' the SQL like query issued to wmi
Dim colServices ' the result of our query as collection
Dim objService ' each individual result 
Dim array1()
Dim array2 ' array2 is created by filter, and sized automatically
Dim a ' counter used for array2 population
Dim b ' counter used for array2 enumeration
Dim i ' counter used for array1
Dim numServices ' used to add 1 to for zero based UBOUND command
Dim numProcesses ' same thing
a = 0
i = 0 
computer = "."
wmiRoot = "winmgmts:\\" & Computer & "\root\cimv2"
Set objWMIService = GetObject(wmiRoot)
wmiQuery = "Select * from Win32_Service Where State <> 'Stopped'"
Set colServices = objWMIService.ExecQuery _
	  (wmiQuery)  
For Each objService In colServices
	ReDim Preserve array1(i)	
    array1(i) = objService.ProcessID 
           i = i + 1
Next






wmiQuery = "Select * from Win32_Service Where ProcessID = '" & _
	            array2(b) & "'"
Set colServices = objWMIService.ExecQuery _
        (wmiQuery)
 	    Wscript.Echo "Process ID: " & array2(b)
	 For Each objService in colServices
	Wscript.Echo VbTab & objService.DisplayName 
	Next




numServices = UBound(array1) + 1 ' due to being zero based
numProcesses = UBound(array2) + 1 ' same reason
WScript.Echo("there are " & numServices & " Services" & _
		" running inside " & numProcesses & " Processes")

