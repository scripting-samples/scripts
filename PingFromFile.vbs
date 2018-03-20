On Error Resume Next

'open the file system object
Set objFSO = CreateObject("Scripting.FileSystemObject")
set WSHShell = wscript.createObject("WScript.shell")
'open the data file
Set objTextStream = objFSO.OpenTextFile("wslist.txt", ForReading)
logfile="results.csv"
'make an array from the data file
RemotePC = Split(oTextStream.ReadAll, vbNewLine)
'close the data file
objTextFile.Close
For Each strComputer In RemotePC

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Notepad.exe'")
   For Each objProcess in colProcessList
       objProcess.Terminate()
   Next

Next
