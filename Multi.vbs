On Error Resume Next
 
'strMachines = "sapcompu,nuchita,cambalache,MP1234"
'aMachines = split(strMachines, ",")
'Dim objFso.objFileHandle, strDisplayString
'Set objFso = 'WScript.CreateObject("Scripting.FileSystemObject")
'Set objFileHandle = objFso.OpenTextFile '("C:\Work\PingTest.txt", 1)
'objFileHandle.Close
Set oFSO = CreateObject("Scripting.FileSystemObject")
set WSHShell = wscript.createObject("wscript.shell")
'open the data file
Set oTextStream = oFSO.OpenTextFile("wslist.txt")
'make an array from the data file
RemotePC = Split(oTextStream.ReadAll, vbNewLine)
'close the data file
oTextStream.Close
For Each strTarget In RemotePC
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ping -n 2 -w 1000 " & strTarget)
strPingResults = LCase(objExec.StdOut.ReadAll)
If InStr(strPingResults, "reply from") Then
    WScript.Echo VbCrLf & strTarget & " responded to ping."
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strTarget & "\root\cimv2")
    Set colCompSystems = objWMIService.ExecQuery("SELECT * FROM " & _
    "Win32_ComputerSystem")
   For Each objCompSystem In colCompSystems
      WScript.Echo "Host Name: " & LCase(objCompSystem.Name)
	
    Next
 Else
    WScript.Echo VbCrLf & strTarget & " did not respond to ping."
  End If

Next