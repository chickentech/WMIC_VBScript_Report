Set objWMI = _
   GetObject("winmgmts:{impersonationLevel=impersonate}\\.\root\cimv2")
Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process")
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objNewFile = objFS.CreateTextFile("writeoutput.txt")
objNewFile.WriteLine "Process Report -- Date: " & Now() & vbCrLf
For Each objProcess In colProcesses
   objNewFile.WriteLine "Process Name: "    & objProcess.Name
   objNewFile.WriteLine "Executable Path: " & objProcess.ExecutablePath
   objNewFile.WriteLine _
        ".............................................."
   objNewFile.WriteLine vbCrLf
Next
objNewFile.Close
