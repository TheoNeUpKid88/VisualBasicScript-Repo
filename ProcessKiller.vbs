' Author: TheOneUpKid
' Purpose: Terminate process that are running or hanging'
' Version 1.7 - March 2017
' ------------------------ -------------------------------' 
Option Explicit
Dim objWMIService, objProcess, colProcess
Dim strComputer, strProcessKill 
strComputer = "."
strProcessKill = InputBox("Name of Program Or Process")

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\cimv2") 

Set colProcess = objWMIService.ExecQuery _
("Select * from Win32_Process Where Name = " & "'"&strProcessKill &"'")
For Each objProcess in colProcess
	On Error Resume Next
	objProcess.Terminate()
	if Err.Number <> 0 Then
		Err.Clear
		On Error GoTo 0
		Exit For
	End If
Next 
WScript.Quit 
