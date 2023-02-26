'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , MS,  3/5/2006
'
' NAME: DisplayAdminTools.vbs
'
' COMMENT: Key concepts are listed below:
'1.Uses the Shell.appliation object to obtain the application object.
'2.Uses the namespace method to connect to a folder by using the special folder constant
'3.Uses the items method to create a collection of items in the special folder. 
'4.Uses for each next to iterate through the colItems collection
'5. Uses wscript.echo to echo out the name of the items in the collection.
'==========================================================================

Set objshell = CreateObject("Shell.Application")
Set objNS = objshell.namespace(&h2f)
Set colitems = objNS.items
For Each objitem In colitems
	WScript.Echo objitem.name
Next