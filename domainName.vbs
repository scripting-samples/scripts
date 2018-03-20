Set WshNtwk = WScript.CreateObject("WScript.Network")  'creates an instance
PropertyInfo = "User Domain" & vbTab & "= " & WshNtwk.UserDomain & _
	vbCrLf & "computer Name" & vbTab & "= " & WshNtwk.ComputerName & _
	vbCrLf & "User Name" & vbTab & "= " & WshNtwk.UserName & vbCrLf
MsgBox PropertyInfo, vbOkOnly , "WshNtwk Properties Example"