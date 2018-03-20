Set WshNtwk = WScript.CreateObject("WScript.Network")
PropertyInfo = "User domain" & vbTab & "= " & WshNtwk.UserDomain & _
vbCrLf & "Computer name" & vbTab & "= " & WshNtwk.ComputerName & _
vbCrLf & "User name" & vbTab & "= " & WshNtwk.UserName & vbCrLf

MsgBox PropertyInfo, vbOkOnly , "WshNtwk Properties Example"

' _characters illustrates a continuation char and is used to indicate a statement is # ' continued on the next line.
' & is a concatenation character