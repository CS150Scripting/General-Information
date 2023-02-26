'==========================================================================
'
' VBScript:  AUTHOR: Ed Wilson , msft,  2/17/2006
'
' NAME: <CreateAddRemoveShortCut.vbs>
'
' COMMENT: Key concepts are listed below:
'1. Uses wscript.shell to create shortcut on the desktop. 
'2. The hard part was passing an argument to the target, which is not allowed
'3. in the target argument. You need to use the arguments property instead.
'==========================================================================
Option Explicit
Dim objShell 'instance of the wshSHell object
Dim strDesktop 'pointer to desktop special folder
Dim objShortCut 'used to set properties of the shortcut. Comes from using createShortCut
Dim strTarget
strTarget = "control.exe"
set objShell = CreateObject("WScript.Shell")
strDesktop = objShell.SpecialFolders("Desktop")

set objShortCut = objShell.CreateShortcut(strDesktop & "\AddRemove.lnk")
objShortCut.TargetPath = strTarget 
objShortCut.Arguments = "appwiz.cpl"
objShortCut.IconLocation = "%SystemRoot%\system32\SHELL32.dll,21"
objShortCut.description = "Add remove Programs"
objShortCut.Save

