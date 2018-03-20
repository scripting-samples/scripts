On Error Resume Next
 
strComputer = "."
arrTargets = Array("sapcompu", "nuchita", "cambalache")
 
 
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
For Each strTarget In arrTargets
 
  Set colPings = objWMIService.ExecQuery _
   ("Select * From Win32_PingStatus where Address = '" & strTarget & "'")
  If Err = 0 Then
    Err.Clear
    For Each objPing in colPings
      If objPing.StatusCode = 0 Then
        Wscript.Echo VbCrLf & strTarget & " responded to ping."
        Wscript.Echo "Responding Address: " & objPing.ProtocolAddress
        Wscript.Echo "Responding Name: " & objPing.ProtocolAddressResolved
        Wscript.Echo "Bytes Sent: " & objPing.BufferSize
        Wscript.Echo "Time: " & objPing.ResponseTime & " ms"
        Wscript.Echo "TTL: " & objPing.ResponseTimeToLive & " seconds"
        GetName
      Else
        WScript.Echo VbCrLf & strTarget & " did not respond to ping."
        WScript.Echo "Status Code: " & objPing.StatusCode
      End If
    Next
  Else
    Err.Clear
    If ExecPing = True Then
      GetName
    End If
  End If
 
Next
 
'******************************************************************************
 
Function ExecPing
 
  Set objShell = CreateObject("WScript.Shell")
  Set objExec = objShell.Exec("ping -n 2 -w 1000 " & strTarget)
  strPingResults = LCase(objExec.StdOut.ReadAll)
  If InStr(strPingResults, "reply from") Then
    WScript.Echo VbCrLf & strTarget & " responded to ping."
    ExecPing = True
  Else
    WScript.Echo VbCrLf & strTarget & " did not respond to ping."
    ExecPing = False
  End If
 
End Function
 
'******************************************************************************
 
Sub GetName
 
  Err.Clear
  Set objWMIServiceRemote = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strTarget & "\root\cimv2")
  If Err = 0 Then
    Err.Clear
    Set colCompSystems = objWMIServiceRemote.ExecQuery("SELECT * FROM " & _
     "Win32_ComputerSystem")
    For Each objCompSystem In colCompSystems
      WScript.Echo "Host Name: " & LCase(objCompSystem.Name)
    Next
  Else
    Err.Clear
    WScript.Echo "Unable to connect to WMI on " & strTarget & "."
  End If
 
End Sub

