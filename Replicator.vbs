Option Explicit
' =====================================================================================================
' Author: Mark Matthias
' Date 1/09/2007
' Name: Replicator.vbs
' Support File: ReplicatorListing.txt

' Summary: 
	' ***** Execute the script with 'cscript Replicator.vbs' at a command line. *****
	' Replicate a set of directories listed in an input file to a destination directory listed in the code as "Destination"
	' Source directory is never changed, but destination will match whatever is in source, i.e. files and directories can be removed.
	' Every run appends progress and errors to a logfile in the local directory.
	' Errors do not stop the script from continuing the backup of everything it can access.
	' Recursive subroutines (ReplicateFolders and ReplicateFiles) courtesy of  Bilal Patel; http://cwashington.netreach.net/depo/view.asp?Index=286
	'  Error Checking and logging was added to the above subroutines.
	
' Notes:
	' 1.  Input file should be local to this script.  Lines can be commented with a hash (#), and paths should not be quoted.
	' 2. First line of input file needs to begin with a plus, and then specify a full path to a DESTINATION.
	' 2. DESTINATION will be used for all source paths after that, until a new DESTINATION is found (a new line with a + at the start)
	' 3.  Below, edit KillOutlookFlag value to be 0 (default) or 1.  Setting it to 1 means Outlook will be killed first. This is  useful when backing up pop server mail.
	' 4. ***** Execute the script with 'cscript Replicator.vbs' at a command line. *****
	
' =====================================================================================================
' =====================================================================================================
' Set KillOutlookFlag to 1 if you need outlook stopped first (to allow pst file backups), or 0 (zero) if not needed.
' This is mainly useful if backing up a mail directory that uses a pop server to download mail to the Outlook directory.
Dim KillOutlookFlag
wscript.echo "Usage: cscript relicator.vbs 1|2 0|1 --> arg1 sets listing file 1|2, arg2 sets outlook kill off|on
KillOutlookFlag = WScript.Arguments.Item (1) ' 0 for do not kill, 1 for kill

' =====================================================================================================
' =====================================================================================================
CheckEngine
Dim fso
Dim MyLog, WshShell, KillCmd, BackupDir, FirstChar, Destination
Dim LastDirectoryName, FinalDestination, oExec, DirectoryListing
Dim InstallDir, DirList, LogFile, MyDate
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const AsASCII = 0, OverWrite = 1
Dim LastChar
MyDate = Now
Set fso = CreateObject("Scripting.FileSystemObject")
InstallDir = fso.GetAbsolutePathName ( "." )
LogFile = fso.BuildPath(InstallDir, "ReplicatorLog.txt")

' read either a 1 or a 2 from command line, to decide which listing to read.
' this does not prevent loading more than one backup destinations and sources that has always been here.
' ReplicatorLog.txt = 1 and ReplicatorLog2.txt = 2
Dim ListSwitch
ListSwitch = WScript.Arguments.Item (0)
If (ListSwitch = 2) Then
	DirList =  fso.BuildPath(InstallDir, "ReplicatorListing2.txt")
Else
	DirList =  fso.BuildPath(InstallDir, "ReplicatorListing1.txt")
End If

' set up log file
If (fso.FileExists(LogFile)) Then
	Set MyLog = fso.OpenTextFile( LogFile, ForAppending, AsASCII )
Else
	Set MyLog = fso.CreateTextFile( LogFile, AsASCII )
End If

MyLog.WriteLine (vbcrlf & "===================================================" & vbcrlf & "Run Date: " & MyDate)
wscript.echo     vbcrlf & "===================================================" & vbcrlf & "Run Date: " & MyDate
LogIt ("INPUT FILE: " & DirList)
LogIt ("LOG FILE: " & LogFile)
MyLog.WriteLine ("")
wscript.echo ("")
' =====================================================================================================
' kill outlook (or not)
If ( KillOutlookFlag = 1 )  Then
	LogIt ("Outlook is set be killed...")
	wscript.sleep 5000
	KilProcess "OUTLOOK.EXE"
	LogIt ("Outlook is killed.")
Else
	LogIt ("Outlook is NOT set to be killed...")
End If
'=====================================================================================================

' open input file
Set DirectoryListing  = fso.OpenTextFile(DirList, ForReading)

Do While Not DirectoryListing.AtEndOfStream
	' verify folder exists and its not commented
	BackupDir = DirectoryListing.ReadLine 
	BackupDir = Trim(BackupDir)
	FirstChar = Left(BackupDir, 1)
	LastChar = Right(BackupDir, 1)
	
	If (LastChar = "\") Then
		BackupDir = Mid ( BackupDir, 1, (Len(Backupdir) - 1) )
	End If
	
	If (FirstChar = "+") Then
		Destination = Mid(BackupDir, 2)
		Destination = Trim(Destination)
		
		If Not fso.FolderExists(Destination) Then
			fso.CreateFolder (Destination)
			' LogIt ("ERROR: Destination " & Destination & " is not valid. Exiting.")
			' wScript.Quit
		End If
		
		LogIt ("DESTINATION: " & Destination)
	ElseIf (fso.FolderExists(BackupDir) And (FirstChar <> "#") ) Then
		LastDirectoryName = Mid(BackupDir, InStrRev(BackupDir, "\") + 1)	
		FinalDestination = fso.BuildPath(Destination, LastDirectoryName)
		
		If ( fso.FolderExists (FinalDestination) = False ) Then
			fso.CreateFolder (FinalDestination)
		End If
		
		LogIt ("COPY: " & BackupDir & " --> " & FinalDestination)
		ReplicateFolders fso, BackupDir, FinalDestination
	End If
Loop

LogIt ("Done" & vbcrlf)

'******************
' Sub ReplicateFolders
'
' This procedure replicates between the source and the destination
' directories at the folder level. A recursive search is done
' between the 2 directories and folders compared. If a particular
' folder on the source does not exist on the destination at any level then the
' source folder and all folders and files associated with it are
' copied to the destination. If a particular folder on the destination
' does not exist on the source at any level then the destination folder
' is removed from the destination directory.
'
'******************


Sub ReplicateFolders (fso, strSourcefolderpath, strDestinationfolderpath)
	Dim aFolderArraySource
	Dim aFolderArrayDestination
	Dim FolderListSource
	Dim FolderListDestination
	Dim oFolderSource
	Dim oFolderDestination
	Dim bSourceExists
	Dim bDestinationExists

	On Error Resume Next
	Err.Clear 
	Set aFolderArraySource = fso.GetFolder(strSourcefolderpath)
	Set aFolderArrayDestination = fso.GetFolder(strDestinationfolderpath)
	Set FolderListSource = aFolderArraySource.SubFolders
	Set FolderListDestination = aFolderArrayDestination.SubFolders

	' Compare to see if destination folder does not exist. If it does not
	' then copy from the source.

	For Each oFolderSource in FolderListSource
	  bDestinationExists = 0
	  For each oFolderDestination in FolderListDestination
	    If oFolderSource.Name = oFolderDestination.Name then
	      bDestinationExists = 1
	      Exit For
	    End If
	  Next
	  If bDestinationExists = 0 then
		' LogIt ("oFolderSource="  & oFolderSource & vbcrlf & "strDestinationfolderpath=" & strDestinationfolderpath & "\")
	    oFolderSource.Copy strDestinationfolderpath & "\" 
		If Err.Number <> 0 Then
			LogIt ("ERROR: File Delete: " & CStr(Err.Number) & " --> " & Err.Description & Err.Source & " " &  strDestinationfolderpath & "\" & oFileDestination.Name )
			Err.Clear
		End If
	  Else
	    'This is the recursive bit. Traverse the path one level down
	    ReplicateFolders fso, strSourcefolderpath & "\" & oFolderSource.Name,_
	strDestinationfolderpath & "\" & oFolderDestination.Name
	  End if
	Next
	' After taking care of the folders, deal with the files at each folder level.
	ReplicateFiles fso, strSourcefolderpath, strDestinationfolderpath


	' Compare to see if a folder on the destination drive does not exist
	' in the source directory. If this is the case then delete the destination
	' folder.

	For Each oFolderDestination in FolderListDestination
	  bSourceExists = 0
	  For each oFolderSource in FolderListSource
	    If oFolderDestination.Name = oFolderSource.Name then
	      bSourceExists = 1
	      Exit For
	    End If
	  Next
	  If bSourceExists = 0 then
	    fso.DeleteFolder strDestinationfolderpath & "\" & oFolderDestination.Name, true
		If Err.Number <> 0 Then
			LogIt ("ERROR: Folder Delete: " & Err.Number & " --> " & Err.Description & " " & strDestinationfolderpath & "\" & oFolderDestination.Name )
			Err.Clear
		End If
	  End if
	Next
End Sub

'******************
' Sub ReplicateFiles
'
' This procedure replicates between the source and the destination
' directories at the file level.
' If a particular file on the source does not exist on the destination
' at any level then the source file is copied to the destination.
' If a particular file on the destination directory
' does not exist on the source at any level then the destination file
' is removed from the destination directory.
'
'******************
Sub ReplicateFiles (fso, strSourcefolderpath, strDestinationfolderpath)
	Dim aFileArraySource
	Dim aFileArrayDestination
	Dim FileListSource
	Dim FileListDestination
	Dim oFileSource
	Dim oFileDestination
	Dim bSourceExists
	Dim bDestinationExists
	Dim SourceDate, DestDate
	On Error Resume Next
	Err.Clear

	Set aFileArraySource = fso.GetFolder(strSourcefolderpath)
	Set aFileArrayDestination = fso.GetFolder(strDestinationfolderpath)
	Set FileListSource = aFileArraySource.Files
	Set FileListDestination = aFileArrayDestination.Files

	' Comparing the array entry properties (name and date last modified) of each array.
	' If the source file array entry matches the destination file array entry then
	' the source file is not copied to the destination directory.
	' Otherwise, the source file is copied to the destination directory and
	' any existing copy of the same file in the destination directory
	' is overwritten.
	
	For each oFileSource in FileListSource
	  bDestinationExists = 0
	  Dim ProblemFile
	  For each oFileDestination in FileListDestination
	    If oFileSource.Name = oFileDestination.Name then
	      If oFileSource.DateLastModified = oFileDestination.DateLastModified then
	        bDestinationExists = 1
	        Exit For
	      End If
	    End If
	  Next
	  If bDestinationExists = 0 then
	    oFileSource.Copy strDestinationfolderpath & "\" & oFileSource.Name
		If Err.Number <> 0 Then
			' LogIt ("WARNING: File Copy: " & Err.Number & " --> " & Err.Description & " " & strDestinationfolderpath & "\" & oFileSource.Name )
			If Not (fso.FileExists(strDestinationfolderpath & "\" & oFileSource.Name)) Then
				' LogIt ("SUCCESS: The file exists in destination Directory")
			' Else
				ProblemFile = fso.BuildPath(strDestinationfolderpath, oFileSource.Name)
				LogIt ("ERROR: File Copy: " & Err.Number & " --> " & Err.Description & " " & strDestinationfolderpath & "\" & oFileSource.Name  & vbcrlf & "   ** AND File does not exist in Destination folder!!")
			End If
			Err.Clear
		End If
	  End If
	Next

	' Comparing the array entry properties (name and date last modified) of each array.
	' If the destination file array entry matches the source file array entry then
	' the destination file is not deleted from the destination directory.
	' Otherwise, the destination file is deleted from the destination directory.
	
	For each oFileDestination in FileListDestination
	  bSourceExists = 0
	  For each oFileSource in FileListSource
	    If oFileDestination.Name = oFileSource.Name then
		  DestDate = convertDate(oFileDestination.DateLastModified) ' drop the time
		  SourceDate = convertDate(oFileSource.DateLastModified)    ' drop the time (Win 8 is weird)
		  If SourceDate = DestDate Then		
	        bSourceExists = 1
	        Exit For
	      End If
	    End If
	  Next
	  If bSourceExists = 0 then
		LogIt ("INFO: Deleting from backup location: " & strDestinationfolderpath & "\" & oFileDestination.Name)
	    fso.DeleteFile strDestinationfolderpath & "\" & oFileDestination.Name, true
		If Err.Number <> 0 Then
			LogIt ("ERROR: File Delete: " & CStr(Err.Number) & " --> " & Err.Description & Err.Source & " " &  strDestinationfolderpath & "\" & oFileDestination.Name )
			Err.Clear
		End If
	  End If
	Next
End Sub

Function convertDate(strDate)
  convertDate = DatePart("d", strDate) & "/" & DatePart("m", strDate) & "/" & DatePart("yyyy", strDate)
  
End Function

Sub KilProcess (strProcessKill)
	Dim objWMIService, objProcess, colProcess
	Dim strComputer, ct
	ct = 0
	strComputer = "."
	' strProcessKill = "'OUTLOOK.exe'" 
	On Error Resume Next
	Err.Clear
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & "'" & strProcessKill & "'")
	If Err.Number <> 0 Then
		Wscript.echo ("ERROR: " & strProcessKill & "was not running")
		LogIt ("ERROR: " & strProcessKill & "was not running")
		Err.Clear
		End If
	Err.Clear 
	For Each objProcess in colProcess
		objProcess.Terminate()
		ct = ct + 1
	Next 	
	If ct = 0 Then
		wscript.echo ("WARNING: " & strProcessKill & " not running.")
		LogIt ("WARNING: " & strProcessKill & " not running.")
	Else
		wscript.echo ("Success: " & strProcessKill & " was stopped.")
		LogIt ("Success: " & strProcessKill & " was stopped.")
	End If
End Sub

Function LogIt (printStr)
	' Dim fso
	' Set fso = CreateObject("Scripting.FileSystemObject")
	MyLog.WriteLine (Time & ": " & printStr)
	wscript.echo Time & ": " & printStr
End Function

Sub CheckEngine
  Dim pcengine
  pcengine = LCase(Mid(WScript.FullName, InstrRev(WScript.FullName,"\")+1))
  If Not pcengine="cscript.exe" Then
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run "CSCRIPT.EXE """ & WScript.ScriptFullName & """"
    WScript.Quit
  End If
End Sub