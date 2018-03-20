On Error Resume Next
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNicConfigs = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
WScript.Echo VbCrLf & "IP Settings"
 
For Each objNicConfig In colNicConfigs
 
  WScript.Echo VbCrLf & "  Network Adapter " & objNicConfig.Index
  WScript.Echo "    " & objNicConfig.Description & VbCrLf
  WScript.Echo "    DHCP Enabled:                   " & _
   objNicConfig.DHCPEnabled
  If Not IsNull(objNicConfig.IPAddress) Then
    strIPAddresses = Join(objNicConfig.IPAddress)
  Else
    strIPAddresses = ""
  End If
  WScript.Echo "    IP Address(es):                 " & strIPAddresses
  If Not IsNull(objNicConfig.IPSubnet) Then
    strIPSubnet = Join(objNicConfig.IPSubnet)
  Else
    strIPSubnet = ""
  End If
  WScript.Echo "    Subnet Mask(s):                 " & strIPSubnet
  If Not IsNull(objNicConfig.DefaultIPGateway) Then
    strDefaultIPGateway = Join(objNicConfig.DefaultIPGateway)
  Else
    strDefaultIPGateway = ""
  End If
  WScript.Echo "    Default Gateways(s):            " & strDefaultIPGateway
  If Not IsNull(objNicConfig.GatewayCostMetric) Then
    strGatewayCostMetric = Join(objNicConfig.GatewayCostMetric)
  Else
    strGatewayCostMetric = ""
  End If
  WScript.Echo "    Gateway Metric(s):              " & strGatewayCostMetric
  WScript.Echo "    Interface Metric:               " & _
   objNicConfig.IPConnectionMetric
  WScript.Echo "    Connection-specific DNS Suffix: " & _
   objNicConfig.DNSDomain
 
Next

