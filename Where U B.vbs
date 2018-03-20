Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
If not objFSO.FileExists("servers.txt") THEN
wscript.echo "Please create a file named 'servers.txt' with one PC name to be pinged per line,"&_
vbcrlf&"with a hard return at the end of each line."
wscript.quit
end if
tempobj="temp.txt"

Set objTextFile = objFSO.OpenTextFile("servers.txt", ForReading)
logfile="results.csv"
Set ofile=objFSO.CreateTextFile(logfile,True)
strText = objTextFile.ReadAll
objTextFile.Close
wscript.echo "Ping batch starting"
ofile.WriteLine ","&"Ping Report -- Date: " & Now() & vbCrLf
arrComputers = Split(strText, vbCrLF)
for each item in arrcomputers
objShell.Run "cmd /c ping -n 1 -w 1000 " & item & " >temp.txt", 0, True
Set tempfile = objFSO.OpenTextFile(tempobj,ForReading)
Do Until tempfile.AtEndOfStream 
temp=tempfile.readall
  striploc = InStr(temp,"[")
        If striploc=0 Then
            strip=""
        Else
            strip=Mid(temp,striploc,16)
            strip=Replace(strip,"[","")
            strip=Replace(strip,"]","")
            strip=Replace(strip,"w"," ")
            strip=Replace(strip," ","")
        End If      
    
        If InStr(temp, "Reply from") Then
             ofile.writeline item & ","&strip&","&"Online." 
        ElseIf InStr(temp, "Request timed out.") Then
             ofile.writeline item &","&strip&","&"No response (Offline)." 
        ELSEIf InStr(temp, "try again") Then
             ofile.writeline item & ","&strip&","&"Unknown host (no DNS entry)."
         
End If 
Loop
Next
tempfile.close
objfso.deletefile(tempobj)
ofile.writeline 
ofile.writeline ","&"Ping batch complete "&now()
wscript.echo "Ping batch completed.  The results will now be displayed."
objShell.Run("""C:\Program Files\Microsoft Office\OFFICE11\excel.exe """&logfile)