Const DriveTypeRemovable = 1
Const DriveTypeFixed = 2
Const DriveTypeNetwork = 3
Const DriveTypeCDROM = 4
Const DriveTypeRAMDisk = 5

Const TempDirName = "\tmp"
Const EnvReg = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"


Set colArgs = WScript.Arguments.Named


TempDriveTypes = Array(DriveTypeRAMDisk, DriveTypeRemovable, DriveTypeFixed)
TempVars = Array("TMP", "TEMP")

Set FSO = CreateObject("Scripting.FileSystemObject")
Set Drives = FSO.Drives

Set WshShell = CreateObject("WScript.Shell")

SystemDrive = WshShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
TempDriveEnv = WshShell.expandEnvironmentStrings("%TEMPDRIVE%")
TempDriveReg = RegGetTempEnv( EnvReg & "TEMPDRIVE" )
WinDir = WshShell.expandEnvironmentStrings("%WINDIR%")

needreboot = false
foundnew = false 

TempDriveTag = TempDriveEnv & "\#tempdrive" 

If TempDriveEnv = SystemDrive OR TempDriveEnv = "%TEMPDRIVE%" OR NOT FSO.FileExists(TempDriveTag) OR TempDriveReg <> TempDriveEnv then
 'We need to find a new TempDrive
  'script.Echo TempDriveTag & " : " & FSO.FileExists(TempDriveEnv & "\#tempdrive")

 For each DriveType in TempDriveTypes
  For each Drive in Drives

   If Drive.DriveType = DriveType Then

    'script.Echo Drive & " : " & FSO.FileExists(Drive & "\#tempdrive")
    If FSO.FileExists(Drive & "\#tempdrive" ) then
     TempDir = Drive & TempDirName
     If CheckTempDir(TempDir) then
      RegSetTempEnv(TempDir)
      foundnew = true
      Exit For
     End If 
    End If

   End If


  Next

  If foundnew = true Then
   Exit For
  End If

 Next

 If foundnew = false Then
  'script.Echo "FoundNew: " & foundnew
  If CheckTempDir( SystemDrive & TempDirName ) Then
   RegSetTempEnv( SystemDrive & TempDirName )
  Else
   RegSetTempEnv( WinDir & "\Temp")
  End If
'  needreboot = true
 End If
End If


If needreboot = true AND colArgs.Exists("r") then
 Reboot
End If




Function CheckTempDir(TmpDir) 'As Boolean
 'script.Echo "CheckTempDir.TmpDir: " & TmpDir
 TmpDrive = Left(TmpDir,2)
 If FSO.FolderExists(TmpDir) AND ISPathWriteAble(TmpDir) then
  CheckTempDir = true
 Else
  'script.Echo "#ERROR: Found " & TmpDrive & " as #tempdrive, but " & TmpDir & " either does not Exist or not Writeable"
  CheckTempDir = false
 End If
End function

Function IsPathWriteable(Path) 'As Boolean
 Set localfso = CreateObject("Scripting.FileSystemObject")
 localtempfile = Path & "\" & localfso.GetTempName() & ".tmp"
 On Error Resume Next
  localfso.CreateTextFile localtempfile
  IsPathWriteable = Err.Number = 0
  localfso.DeleteFile localtempfile
 On Error Goto 0
End Function

Function RegSetTempEnv(TmpDir) 'As Boolean
 'script.Echo "RegSetTempEnv.TmpDir: " & TmpDir
 TmpDrive=Left(TmpDir,2)
 On Error Resume Next
 oldTmpDrive = WshShell.RegRead( EnvReg & "TEMPDRIVE" )
 If oldTmpDrive <> TmpDrive Then
  WshShell.RegWrite EnvReg & "TEMPDRIVE",TmpDrive,"REG_EXPAND_SZ"
  needreboot = true
 End If

 'script.Echo "RegSetTempEnv.TmpDir.Right: " & Right(TmpDir,Len(TmpDir)-2)
  If Right(TmpDir,Len(TmpDir)-2) = TempDirName Then
   'script.Echo "RegSetTempEnv: If Right EQ"
   SetTempKeys("%TEMPDRIVE%" & TempDirName)
  Else
   'script.Echo "RegSetTempEnv: If Right ELSE"
   SetTempKeys(TmpDir)
  End If
End Function

Function RegGetTempEnv(EnvName)
 On Error Resume Next
 RegGetTempEnv =  WshShell.RegRead(EnvName) = ""
End Function

Function SetTempKeys(TmpDir)
 For each TempVar in TempVars
  On Error Resume Next
  oldTempKey = WshShell.RegRead(EnvReg & TempVar)
  'script.Echo "SetTempKeys: " & TempVar & ": '" & oldTempKey & "', TmpDir: '" & TmpDir & "'"
  if oldTempKey <> TmpDir then
   'script.Echo "SetTempKeys: oldTempKey not EQ"
   WshShell.RegWrite EnvReg & TempVar,TmpDir,"REG_EXPAND_SZ"
   needreboot = true
  End If
 Next
End Function

Function Reboot

 If colArgs.Exists("f") Then
  rebootforce = " -f "
 Else
  rebootforce = ""
 End If

 RunString = "%comspec% /c %WINDIR%\system32\shutdown.exe -r -t 0 " & rebootforce & " -c Environment_variables_TMP_and_TEMP_updated._Rebooting..."

'script.Echo "Reboot: " & RunString
 Return = WshShell.Run(RunString, 1, TRUE)
End Function
