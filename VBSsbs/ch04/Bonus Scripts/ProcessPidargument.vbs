'==========================================================================
'
'
' NAME: <ProcessPidArgument.vbs>
'
' AUTHOR: Ed Wilson , MS
' DATE  : 10/28/2003
'
' COMMENT: <Uses win32_Process to display a lot of information about
' processes on the machine. >
'1. gets command line argument for PID to investigate
'2. Uses vwhere to hold a variable where clause
'==========================================================================

Option Explicit 
On Error Resume Next
dim wmiQuery
dim objWMIService
dim colItems
dim objItem
Dim vwhere, argPID
argPID = WScript.Arguments.Item(0)
vWhere = "where ProcessID =" & argPID
wmiQuery = "Select * from win32_Process " & vwhere
Set objWMIService = GetObject("winmgmts:\\")
Set colItems = objWMIService.ExecQuery(wmiQuery)

For Each objItem in colItems
     WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "CommandLine: " & objItem.CommandLine
      WScript.Echo "CreationClassName: " & objItem.CreationClassName
      WScript.Echo "CreationDate: " & objItem.CreationDate
      WScript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
      WScript.Echo "CSName: " & objItem.CSName
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "ExecutablePath: " & objItem.ExecutablePath
      WScript.Echo "ExecutionState: " & objItem.ExecutionState
      WScript.Echo "Handle: " & objItem.Handle
      WScript.Echo "HandleCount: " & objItem.HandleCount
      WScript.Echo "InstallDate: " & objItem.InstallDate
      WScript.Echo "KernelModeTime: " & objItem.KernelModeTime
      WScript.Echo "MaximumWorkingSetSize: " & objItem.MaximumWorkingSetSize
      WScript.Echo "MinimumWorkingSetSize: " & objItem.MinimumWorkingSetSize
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "OSCreationClassName: " & objItem.OSCreationClassName
      WScript.Echo "OSName: " & objItem.OSName
      WScript.Echo "OtherOperationCount: " & objItem.OtherOperationCount
      WScript.Echo "OtherTransferCount: " & objItem.OtherTransferCount
      WScript.Echo "PageFaults: " & objItem.PageFaults
      WScript.Echo "PageFileUsage: " & objItem.PageFileUsage
      WScript.Echo "ParentProcessId: " & objItem.ParentProcessId
      WScript.Echo "PeakPageFileUsage: " & objItem.PeakPageFileUsage
      WScript.Echo "PeakVirtualSize: " & objItem.PeakVirtualSize
      WScript.Echo "PeakWorkingSetSize: " & objItem.PeakWorkingSetSize
      WScript.Echo "Priority: " & objItem.Priority
      WScript.Echo "PrivatePageCount: " & objItem.PrivatePageCount
      WScript.Echo "ProcessId: " & objItem.ProcessId
      WScript.Echo "QuotaNonPagedPoolUsage: " & objItem.QuotaNonPagedPoolUsage
      WScript.Echo "QuotaPagedPoolUsage: " & objItem.QuotaPagedPoolUsage
      WScript.Echo "QuotaPeakNonPagedPoolUsage: " & objItem.QuotaPeakNonPagedPoolUsage
      WScript.Echo "QuotaPeakPagedPoolUsage: " & objItem.QuotaPeakPagedPoolUsage
      WScript.Echo "ReadOperationCount: " & objItem.ReadOperationCount
      WScript.Echo "ReadTransferCount: " & objItem.ReadTransferCount
      WScript.Echo "SessionId: " & objItem.SessionId
      WScript.Echo "Status: " & objItem.Status
      WScript.Echo "TerminationDate: " & objItem.TerminationDate
      WScript.Echo "ThreadCount: " & objItem.ThreadCount
      WScript.Echo "UserModeTime: " & objItem.UserModeTime
      WScript.Echo "VirtualSize: " & objItem.VirtualSize
      WScript.Echo "WindowsVersion: " & objItem.WindowsVersion
      WScript.Echo "WorkingSetSize: " & objItem.WorkingSetSize
      WScript.Echo "WriteOperationCount: " & objItem.WriteOperationCount
      WScript.Echo "WriteTransferCount: " & objItem.WriteTransferCount
      WScript.Echo " *********************************"

Next