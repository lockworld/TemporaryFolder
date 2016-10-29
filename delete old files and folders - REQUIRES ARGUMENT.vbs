Option Explicit


'wscript.Echo ("Cleaning up temporary Internet files...")
'WScript.Echo ("From " + WScript.Arguments.Item(0))
if (WScript.Arguments.Item(1)>"") then
	'WScript.Echo ("If Older Than " & WScript.Arguments.Item(1))
end if

'---------------------------------------------------
'DECLARE VARIABLES
'---------------------------------------------------


Dim objFSO, objFolder, objShell, objLogFile, objFile
Dim logDirectory, logFile, logText
Dim ofolder, objStream, objNet
Dim strComputer, SPath, FSO, MyFolder
Dim dtstr, KeepForDays


'---------
'GET PARAMETERS
'---------
SPath = WScript.Arguments.Item(0)
if (WScript.Arguments.Item(1)>0) then
	KeepForDays = WScript.Arguments.Item(1)
else
	KeepForDays = 7
end if


dtstr = year(date)

if (month(date)<10) then
    dtstr = dtstr & "0" & month(date)
else
    dtstr = dtstr & month(date)
end if

if (day(date)<10) then
    dtstr = dtstr & "0" & day(date)
else
    dtstr = dtstr & day(date)
end if

if (hour(Now())<10) then
    dtstr = dtstr & "-0" & hour(Now())
else
    dtstr = dtstr & "-" & hour(Now())
end if

if (minute(Now())<10) then
    dtstr = dtstr & "0" & minute(Now())
else
    dtstr = dtstr & minute(Now())
end if

if (second(Now())<10) then
    dtstr = dtstr & "0" & second(Now())
else
    dtstr = dtstr & second(Now())
end if


logDirectory = SPath & "Logs\"
logFile = "\delete old files and folders LOG-" & dtstr & ".txt"

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")


'---------------------------------------------------
'Open Log File:
'  This procedure opens a log file (txt document) in Append-only mode.
'---------------------------------------------------


  ' Check that the logDirectory folder exists
  ' If it doesn't exist, create it
    If objFSO.FolderExists(logDirectory) Then
       Set objFolder = objFSO.GetFolder(logDirectory)
    Else
       Set objFolder = objFSO.CreateFolder(logDirectory)
       'WScript.Echo "Just created " & logDirectory
    End If
  
  ' Check that the logFile file exists
  ' If it doesn't exist, create it
    If objFSO.FileExists(logDirectory & logFile) Then
       Set objFolder = objFSO.GetFolder(logDirectory)
    Else
       Set objFile = objFSO.CreateTextFile(logDirectory & logFile)
       'Wscript.Echo "Just created " & logDirectory & logFile
    End If 


  set objFile = nothing
  set objFolder = nothing

  ' OpenTextFile Method needs a Const value for "OpenMethod"
  ' ForAppending = 8 ForReading = 1, ForWriting = 2
    Const OpenMethod = 2

  Set objLogFile = objFSO.OpenTextFile _
  (logDirectory & logFile, OpenMethod, True)

  objLogFile.WriteLine("--------------------" & vbNewLine & "Cleaning Temporary Files: " & FormatDateTime(Now(),0) & vbNewLine)
  objLogFile.WriteLine("From " & SPath & vbNewLine)
  objLogFile.WriteLine("Deleting files older than " & KeepForDays & " days old" & vbNewLine)
  objLogFile.WriteLine("   " & Date & vbNewLine)

'---------------------------------------------------
'END Log File Procedure
'---------------------------------------------------


'---------------------------------------------------
'FUNCTIONS
'---------------------------------------------------
Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function




'---------------------------------------------------
'Basic delete objects if older than X days procedure
'---------------------------------------------------
'  strComputer = "."
'
'  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
'
'  Set colFiles = objWMIService.ExecQuery _
'      ("ASSOCIATORS OF {Win32_Directory.Name='H:\Temp\' Where " _
'          & "ResultClass = CIM_DataFile")

'  strCurrentDate = Now

'  noDays=CInt(168)

'  For Each objFile In colFiles
'      strFileDate = WMIDateStringToDate(objFile.CreationDate)
'      intHours = DateDiff("h", strFileDate, strCurrentDate)
        'Check if file is more than 168 hours (7 days) old...
'      If intHours >= 168 Then
          'Wscript.Echo (If you want a message box telling you how old the file is)
'          objLogFile.WriteLine(vbNewLine & "  Deleted " & round((Wscript.Echo/24),0) & " day old file: "  & objFile.path & vbNewLine & "--" & vbNewLine)
'          objFile.Delete
'      End If
'  Next



'---------------------------------------------------
'Remove old FOLDERS
'---------------------------------------------------

  Set objShell = CreateObject("WScript.Shell")
  Set objFSO = CreateObject("scripting.filesystemobject")
  Set objNet = CreateObject("WScript.NetWork")
  Set FSO = CreateObject("Scripting.FileSystemObject")

  


  Set MyFolder = FSO.GetFolder(SPath)

  DelSubfolders FSO.GetFolder(SPath),KeepForDays
  

  Sub DelSubFolders(Folder,noDays)
    Dim Datedifference
    Dim MySubFolders
    Dim MyFiles
    Dim MyFile
    Dim MyFolder
  
    '=== Get the collection of Folders in this folder
    Set MySubFolders = Folder.SubFolders 
    
    '=== Get the collection of Files in this folder
    Set MyFiles = Folder.files
  
    '=== If there are subfolders, process them first.
    IF MySubFolders.Count <> 0 Then
      For each MyFolder in MySubFolders
        DelSubFolders MyFolder, noDays
      Next
    End If
  
    '=== If this folder isn't emtpy, process each file to see if they are older than the maximum age limit (noDays).
    IF MyFiles.Count <> 0 Then
      For Each MyFile in MyFiles
        '=== Find out how old the file is compared to current date
        Datedifference = DateDiff("D",MyFile.DateLastModified,Date)
        IF (Datedifference > CInt(noDays)) Then
          '=== This file is old, delete and add entry to logfile
          'outputString = "Deleted: " & MyFile.path
          'WScript.Echo outputString
          'fsOut.WriteLine outputString
          '=== If you just want to do a dry run of this script, comment out
          ' the next line to prevent the file from being deleted. The true
          ' after the delete statement is necessary to force delete

          objLogFile.WriteLine(vbNewLine & "     Deleted File: " & MyFile.path & vbNewLine & "--" & vbNewLine)
          MyFile.delete true
        Else
          'Wscript.Echo MyFile.path & " is OK!"
        End If
      Next
    End If
  
    '=== If this folder is emtpy, meaning no subfolders or files, check if this folder is older than maximum age limit and delete accordingly
    If MySubFolders.Count = 0 And MyFiles.Count = 0 Then
      Datedifference = DateDiff("D",Folder.DateLastModified,Date)
      'NOTE: If this folder is empty, there is no reason to check to see when it was modified
      'IF (Datedifference > CInt(noDays)) Then
        '=== This Folder is old, delete and add entry to logfile
        'outputString = "Deleted: " & Folder.path
        'WScript.Echo outputString
        'fsOut.WriteLine outputString
        '=== If you just want to do a dry run of this script, comment out
        ' the next line to prevent the file from being deleted. The true
        ' after the delete statement is necessary to force delete

        If (Folder.path<>"H:\Temp") Then
            objLogFile.WriteLine(vbNewLine & "     Deleted Folder: " & Folder.path & vbNewLine & "--" & vbNewLine)
     	      Folder.delete true
        End If
        Exit Sub
      'Else
        'Wscript.Echo Folder.path & " is OK!"
      'End If
    End If
  











End Sub




'---------------------------------------------------
'VERIFY Temporary Files folder still exists.
'If it was accidentally deleted, create a new one
'---------------------------------------------------
If objFSO.FolderExists(SPath) Then
       objLogFile.WriteLine(vbNewLine & "   --" & "   FOLDER " & SPath & " CHECKED SUCCESSFULLY.")
       'Wscript.Echo "SUCCESS!"
    Else
       Set objFolder = objFSO.CreateFolder(SPath)
       objLogFile.WriteLine(vbNewLine & "   --" & "   ERROR: PROCESS ERASED 'Temporary Files' folder. A New Temporary Files folder has been created.")
       Wscript.Echo SPath & " has been re-created!"
    End If




objLogFile.WriteLine(vbNewLine & "   DELETE OLD FILES AND FOLDERS HAS COMPLETED SUCCESSFULLY" & vbNewLine & "----------")


