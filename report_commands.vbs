' Computer Report Generator v2 written in VBScript
' Original script written in Powershell
' by: Nathan Behe
' for the Commonwealth of PA - Office of Administration Help Desk
' Created on: 2/12/2015
'

Set objWMI = _
GetObject("winmgmts:{impersonationLevel=impersonate}\\.\root\cimv2")
Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process")
' Open HTML File for Editing
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objNewFile = objFS.CreateTextFile("writeoutput.html")
objNewFile.WriteLine "<html>" + vbCrLf
objNewFile.WriteLine "<head>" + vbCrLf
' Add JQuery
objNewFile.WriteLine "<script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js'></script>" + vbCrLf
objNewFile.WriteLine "<title>Computer Report</title>" + vbCrLf
objNewFile.WriteLine "</head>" + vbCrLf
objNewFile.WriteLine "<body>" + vbCrLf
For Each objProcess In colProcesses
objNewFile.WriteLine "Process Name: "    & objProcess.Name
objNewFile.WriteLine "Executable Path: " & objProcess.ExecutablePath
objNewFile.WriteLine _
".............................................."
objNewFile.WriteLine vbCrLf
Next
objNewFile.WriteLine "</body></html>"
objNewFile.Close
