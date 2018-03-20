On Error Resume Next
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNics = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_NetworkAdapter")
 
WScript.Echo VbCrLf & "Network Adapter Settings"
 
For Each objNic In colNics
 
  WScript.Echo VbCrLf & "  Network Adapter (Device ID)" & _
   objNic.DeviceID
  Wscript.Echo "    Index: " & objNic.Index
  Wscript.Echo "    MAC Address: " & objNic.MACAddress
  Wscript.Echo "    Adapter Type: " & objNic.AdapterType
  Wscript.Echo "    Adapter Type Id: " & objNic.AdapterTypeID
  Wscript.Echo "    Description: " & objNic.Description
  Wscript.Echo "    Manufacturer: " & objNic.Manufacturer
  Wscript.Echo "    Name: " & objNic.Name
  Wscript.Echo "    Product Name: " & objNic.ProductName
  Wscript.Echo "    Net Connection ID: " & objNic.NetConnectionID
  Wscript.Echo "    Net Connection Status: " & objNic.NetConnectionStatus
  Wscript.Echo "    PNP Device ID: " & objNic.PNPDeviceID
  Wscript.Echo "    Service Name: " & objNic.ServiceName
  If Not IsNull(objNic.NetworkAddresses) Then
    strNetworkAddresses = Join(objNic.NetworkAddresses)
  Else
    strNetworkAddresses = ""
  End If
  Wscript.Echo "    NetworkAddresses: " & strNetworkAddresses
  Wscript.Echo "    Permanent Address: " & objNic.PermanentAddress
  Wscript.Echo "    AutoSense: " & objNic.AutoSense
  Wscript.Echo "    Maximum Number Controlled: " & objNic.MaxNumberControlled
  Wscript.Echo "    Speed: " & objNic.Speed
  Wscript.Echo "    Maximum Speed: " & objNic.MaxSpeed
 
Next

