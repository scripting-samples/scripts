Set objWMIService = GetObject("winmgmts:")
Set colNicConfig = objWMIService.ExecQuery("SELECT * FROM " & _
 "Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each objNicConfig In colNicConfig
  WScript.Echo "Network Adapter: " & objNicConfig.Index
  WScript.Echo "  IP Address(es):"
  If Not IsNull(objNicConfig.IPAddress) Then
    For Each strIPAddress In objNicConfig.IPAddress
      WScript.Echo "    " & strIPAddress
    Next
  End If
Next

