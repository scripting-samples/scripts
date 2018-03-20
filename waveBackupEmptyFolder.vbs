Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO1 = CreateObject("Scripting.FileSystemObject")
NumberOfDays = 5 'anything older than 5 days will be deleted

Set objFolder = objFSO.GetFolder("C:\Program Files (x86)\GE Healthcare\UNICORN\UNICORN Database\Backup")
Set objFolder1 = objFSO1.GetFolder("L:\Research and Development\RDD Instrument Data\CMoB\CMoB Instruments\Wave")

For Each aFile in objFolder.Files
  If DateDiff("d", objFile.DateCreated,Now) > NumberOfDays Then
        objFile.Delete True
        End If
  
Next

For Each aFile in objFolder1.Files
  If DateDiff("d", objFile.DateCreated,Now) > NumberOfDays Then
        objFile.Delete True
        End If
Next


