'==========================================================================
'
' NAME: ListFolderSizes.vbs
'
' AUTHOR: ed wilson , mred
' DATE  : 4/6/2006
'
' COMMENT: <Displays size of parent folder, and associated subFolders.>
'1.Uses FileSystemObject and the getFolder method.
'2.Uses subFolders method to get collection of subFolders
'3.Uses FormatNumber function to add comma's to numbers displayed.
'==========================================================================
Option Explicit 
'On Error Resume Next
Dim objFSO 			'the fileSystemObject
Dim objFolder 	'folder object
Dim strFolder		'individual folder form collection
Dim colFolders	'collection of subFolders
Dim strHeader		'header used for reporting

Const noDecimal = 0 'number of decimal places for FormatNumber
strFolder = "c:\windows" 			'Path to specific folder 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolder)
	strHeader= objFolder.Path & vbTab & formatNumber(objFolder.size,noDecimal)
	WScript.echo funline(strHeader)
Set colFolders = objFolder.SubFolders

For Each strFolder In colFolders
	WScript.Echo strFolder.path, formatNumber(strFolder.size,noDecimal)
Next

'*** Function is below ***
Function funline(strIn)
funline = Len(strIN)+1
funline = strIN & VbCrLf & String(funLine,"=")
End Function