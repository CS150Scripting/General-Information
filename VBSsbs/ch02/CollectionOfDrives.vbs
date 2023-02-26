Option Explicit    	 
'On Error Resume Next 
Dim colDrives 		'The collection that comes from WMI
Dim drive     		'An individual drive in the collection
Const DriveType = 3 	'Local drives. From the SDK

set colDrives =_ 
GetObject("winmgmts:").ExecQuery("select size,freespace " &_
	 "from Win32_LogicalDisk where DriveType =" & DriveType)

For Each drive in colDrives 'Walks through the collection.
	WScript.Echo "Drive: " & drive.DeviceID
	WScript.Echo "Size: " & drive.size
	WScript.Echo "Freespace: " & drive.freespace
Next
