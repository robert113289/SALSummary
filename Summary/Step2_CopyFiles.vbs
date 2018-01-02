filename = "X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\StoreList.txt" 


Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)

incomingDir="X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\Incoming\"
localDir="X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\Status\"

Do Until f.AtEndOfStream
  
  strLine = f.ReadLine
   'WScript.Echo strLine
   Lines =len(strLine)
   'msgbox(Lines)
   
   If Lines = 5 Then
   
   strName = "ST" & strLine & ".IN." & "Store_POSInfo.txt"
   
   'msgbox(strName)
   
   Elseif Lines = 4 Then
   
   strName = "ST0" & strLine & ".CO." & "Store_POSInfo.txt"
   
   'msgbox(strName)
   
   Elseif Lines = 3 Then
   
   strName = "ST00" & strLine & ".CO." & "Store_POSInfo.txt"
   
   Else
   
   msgbox("Please check files.")
   
   End If



Set objFSO = CreateObject("Scripting.FileSystemObject")
'dtmValue = Now()
'strFolder = localDir & Month(dtmValue) & "-" & Day(dtmValue) & "-" & Year(dtmValue)

curMonth=right("00" & month(now),2)
curDay=right("00" & day(now),2)
curYear=right("0000" & year(now),4)

strFolder = localDir & curYear & curMonth & curDay

If objFSO.FolderExists(strFolder) then

Else 

objFSO.CreateFolder(strFolder)

End If

   
   
DestinationFile = strFolder &"\"
SourceFile = incomingDir & strName


'MsgBox(SourceFile)
'MsgBox(DestinationFile)


Set fso=CreateObject("Scripting.FilesystemObject")
fso.CopyFile SourceFile, DestinationFile
 
  
Loop
MsgBox("Operation Completed")
f.Close


