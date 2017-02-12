' defrag_all.vbs
' Defrag All Removable and Fixed Drives

Const DriveTypeRemovable = 1
Const DriveTypeFixed = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Drives = objFSO.Drives
Set WshShell = WScript.CreateObject("WScript.Shell")

Call SetLocale(1033)

Minutes = 10
Set colArgs = WScript.Arguments.Named
If colArgs.Exists("minutes") Then
 Minutes = CLng(colArgs.Item("minutes"))
End If

RemovableMultiplier = 6
If colArgs.Exists("removablemultiplier") Then
 RemovableMultiplier = CLng(colArgs.Item("removablemultiplier"))
End If


LogDir = SetupLogTarget()

For each Drive in Drives
 if Drive.DriveType = 1 then
  RunCmd Drive,Minutes * RemovableMultiplier
 end if
Next

For each Drive in Drives
 If Drive.DriveType = 2 Then
  RunCmd Drive,Minutes
 end if
Next

Function RunCmd(DriveString,Minutes)

 LogFile = LogDir & "\" & "defrag_" & Left(DriveString,1) & ".log"

 FileDate = Now()
 firstrun = true

 If objFSO.FileExists(LogFile) Then
  set objLogFile = objFSO.GetFile(LogFile)
  FileDate = CDATE(objLogFile.DateLastModified)
  firstrun = false
 End If

 If CLng(DateDiff("n", FileDate, Now() )) > Minutes OR firstrun = true Then
  RunString = "%comspec% /c echo " & WeekDayName(WeekDay(Now), True) & " " & Now & " - Defragment - " & DriveString
  Return = WshShell.Run(RunString & " >> " & LogFile & " 2>&1", 0, TRUE)

  RunString = "%comspec% /c %WINDIR%\system32\defrag.exe " & DriveString & " -f -v"
  Return = WshShell.Run(RunString & " >> " & LogFile & " 2>&1", 0, TRUE)
 End If
 RunCmd = Return

End Function



Function IsPathWriteable(Path) 'As Boolean

 If Path <> "" Then
  Set localfso = CreateObject("Scripting.FileSystemObject")
  localtempfile = Path & "\" & localfso.GetTempName() & ".tmp"
  On Error Resume Next
   localfso.CreateTextFile localtempfile
   IsPathWriteable = Err.Number = 0
   localfso.DeleteFile localtempfile
  On Error Goto 0
 Else
  IsPathWriteable = 1
 End If

End Function


Function SetupLogTarget

 If colArgs.Exists("logdir") Then
  If IsPathWriteAble(colArgs.Item("logdir")) Then
   SetupLogTarget = colArgs.Item("logdir")
   Exit Function
  End If
 End If

 If WshShell.ExpandEnvironmentStrings("%LOGDIR%") <> "%LOGDIR%" Then
  If IsPathWriteable(WshShell.ExpandEnvironmentStrings("%LOGDIR%")) Then
   SetupLogTarget = WshShell.ExpandEnvironmentStrings("%LOGDIR%")
   Exit Function
  End If
 End If

 If WshShell.ExpandEnvironmentStrings("%TEMP%") <> "%TEMP%" Then
  If IspathWriteable( WshShell.ExpandEnvironmentStrings("%TEMP%") ) Then
   SetupLogTarget = WshShell.ExpandEnvironmentStrings("%TEMP%")
   Exit Function
  End If
 End If

 If WshShell.ExpandEnvironmentStrings("%WINDIR%") <> "%WINDIR%" Then
  If IspathWriteable( WshShell.ExpandEnvironmentStrings("%WINDIR%" & "\Temp" ) ) Then
   SetupLogTarget = WshShell.ExpandEnvironmentStrings("%WINDIR%") & "\Temp"
   Exit Function
  End If
 End If

 Wscript.Echo "#ERROR: I need a place to write logfiles to, but no /logdir: was specified!"

End Function
