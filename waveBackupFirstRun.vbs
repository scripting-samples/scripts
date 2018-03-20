Dim objFSO
Dim sSourceFolder
Dim sDestFolder
Dim sDBFile
Const OVER_WRITE_FILES = True
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
sSourceFolder = "C:\Program Files (x86)\GE Healthcare\UNICORN\UNICORN Database\Backup"
sBackupFolder = "L:\Research And Development\RDD Instrument Data\CMoB\CMoB Instruments\Wave"
sDBFile = "Test.mdb"
 
'If the backup folder doesn't exist, create it.
If Not objFSO.FolderExists(sBackupFolder) Then
    objFSO.CreateFolder(sBackupFolder)
End If
 
'Copy the file as long as the file can be found
If objFSO.FileExists(sSourceFolder & "\" & sDBFile) Then
    objFSO.CopyFile sSourceFolder & "\" & sDBFile, sBackupFolder & "\" & sDBFile, OVER_WRITE_FILES
End if
 
Set objFSO = Nothing