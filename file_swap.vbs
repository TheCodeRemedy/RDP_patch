' User configurable variables
Dim LogFileName, LogFilePath, LogFile, DesiredFilePath
LogFileName = "FileCopy.log"
LogFilePath = "C:\Windows\Logs\TheCodeRemedy\" ' Must include trailing backslash
LogFile = LogFilePath + LogFileName
DesiredFilePath = "C:\TEST" ' Local or UNC path for the desired FileSystemObject, this expects the files to be placed into respective 32 and 64 folders. eg. C:\TEST\32 and C:\TEST\64

' Desired end result version of the files
DesiredMSTSCVersion = "10.0.15063.674"
DesiredMSTSCAXVersion = "10.0.15063.726"
DesiredMSTSCVersion64 = "10.0.15063.674"
DesiredMSTSCAXVersion64 = "10.0.15063.726"

' Nothing configurable past this point
' Declare system variables
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")

' Declare version variables
Dim CurrentMSTSCVersion, CurrentMSTSCAXVersion, CurrentMSTSCVersion64, CurrentMSTSCAXVersion64, TransitionMSTSCVersion, TransitionMSTSCAXVersion, DesiredMSTSCVersion, DesiredMSTSCAXVersion



' Check for existence of, and open log file, create if not found
if not (objFSO.FileExists(LogFile)) then
  'Check for existence of log file path, create otherwise
  if not (objFSO.FolderExists(LogFilePath)) then
    objFSO.CreateFolder(LogFilePath)
  end if

  Set objLogFile = objFSO.CreateTextFile(LogFile,true)
else
  Set objLogFile = objFSO.OpenTextFile(LogFile,2)
end if


' Work
' Start log file
objLogFile.WriteLine("Initializing...")

' Check whether desired files are accessible
if (objFSO.FolderExists(DesiredFilePath)) and (objFSO.FileExists(DesiredFilePath + "\32\mstsc.exe")) and (objFSO.FileExists(DesiredFilePath + "\32\mstscax.dll")) and (objFSO.FileExists(DesiredFilePath + "\64\mstsc.exe")) and (objFSO.FileExists(DesiredFilePath + "\64\mstscax.dll")) then

  ' Ensure backup directory structure exists
  ' Check for existence of C:\TEMP\32
  if not (objFSO.FolderExists("C:\TEMP\32")) then
    if not (objFSO.FolderExists("C:\TEMP")) then
    objLogFile.WriteLine("Directory: C:\TEMP not found, creating")
      objFSO.CreateFolder("C:\TEMP")
    end if
    objLogFile.WriteLine("Directory: C:\TEMP\32 not found, creating")
    objFSO.CreateFolder("C:\TEMP\32")
  end if
  ' and now c:\TEMP\64
  if not (objFSO.FolderExists("C:\TEMP\64")) then
    if not (objFSO.FolderExists("C:\TEMP")) then
    objLogFile.WriteLine("Directory: C:\TEMP not found, creating")
      objFSO.CreateFolder("C:\TEMP")
    end if
    objLogFile.WriteLine("Directory: C:\TEMP\64 not found, creating")
    objFSO.CreateFolder("C:\TEMP\64")
  end if

  ' Retrieve the current version of the files in place
  ' First from System32
  if (objFSO.FolderExists("C:\Windows\System32")) and (objFSO.FileExists("c:\Windows\System32\mstsc.exe")) and (objFSO.FileExists("c:\Windows\System32\mstscax.dll")) then
      CurrentMSTSCVersion = objFSO.GetFileVersion("c:\windows\system32\mstsc.exe")
      objLogFile.WriteLine("INFO: Retrieved file version for C:\Windows\System32\mstsc.exe as " + CurrentMSTSCVersion)
      CurrentMSTSCAXVersion = objFSO.GetFileVersion("c:\windows\system32\mstscax.dll")
      objLogFile.WriteLine("INFO: Retrieved file version for C:\Windows\System32\mstscax.dll as " + CurrentMSTSCAXVersion)
  else
    if (objFSO.FolderExists("C:\Windows\System32")) then
      objLogFile.WriteLine("ERROR: Directory C:\Windows\System32 not found")
    end if
    if (objFSO.FileExists("c:\Windows\System32\mstsc.exe")) then
      objLogFile.WriteLine("ERROR: File C:\Windows\System32\mstsc.exe not found")
    end if
    if (objFSO.FileExists("c:\Windows\System32\mstscax.dll")) then
      objLogFile.WriteLine("ERROR: File C:\Windows\System32\mstscax.dll not found")
    end if
  end if

  ' Now from SysWOW64
  if (objFSO.FolderExists("C:\Windows\SysWOW64")) and (objFSO.FileExists("C:\Windows\SysWOW64\mstsc.exe")) and (objFSO.FileExists("C:\Windows\SysWOW64\mstscax.dll")) then
      CurrentMSTSCVersion64 = objFSO.GetFileVersion("C:\Windows\SysWOW64\mstsc.exe")
      objLogFile.WriteLine("INFO: Retrieved file version for C:\Windows\SysWOW64\mstsc.exe as " + CurrentMSTSCVersion64)
      CurrentMSTSCAXVersion64 = objFSO.GetFileVersion("C:\Windows\SysWOW64\mstscax.dll")
      objLogFile.WriteLine("INFO: Retrieved file version for C:\Windows\SysWOW64\mstscax.dll as " + CurrentMSTSCAXVersion64)
  else
    if (objFSO.FolderExists("C:\Windows\SysWOW64")) then
      objLogFile.WriteLine("ERROR: Directory C:\Windows\SysWOW64 not found")
    end if
    if (objFSO.FileExists("c:\Windows\SysWOW64\mstsc.exe")) then
      objLogFile.WriteLine("ERROR: File C:\Windows\SysWOW64\mstsc.exe not found")
    end if
    if (objFSO.FileExists("c:\Windows\SysWOW64\mstscax.dll")) then
      objLogFile.WriteLine("ERROR: File C:\Windows\SysWOW64\mstscax.dll not found")
    end if
  end if

  ' If the current version of the two files are not identical to desired, take action
  ' First for System32
  if (CurrentMSTSCVersion <> DesiredMSTSCVersion) then
    objLogFile.WriteLine("INFO: C:\Windows\System32\mstsc.exe is not the desired version")
    ' Check if backup already exists, skip if true
    if not (objFSO.FileExists("C:\TEMP\32\mstsc.exe")) then
      objLogFile.WriteLine("INFO: Backing up C:\Windows\System32\mstsc.exe before replacing.")
      objFSO.CopyFile "C:\Windows\System32\mstsc.exe", "C:\TEMP\32\mstsc.exe", false
    else
      objLogFile.WriteLine("INFO: C:\Windows\System32\mstsc.exe backup already exists.")
    end if
    ' take ownership of file in question
    objShell.Run("takeown /f C:\Windows\System32\mstsc.exe"),0,true
    objShell.Run("icacls C:\Windows\System32\mstsc.exe /grant administrators:F"),0,true
    ' replace file
    objLogFile.WriteLine("INFO: Copying desired version of C:\Windows\System32\mstscax.dll.")
    objFSO.CopyFile DesiredFilePath + "\32\mstsc.exe", "C:\Windows\System32\mstsc.exe", true
  else
    objLogFile.WriteLine("INFO: C:\Windows\System32\mstsc.exe is the desired version")
  end if

  if (CurrentMSTSCAXVersion <> DesiredMSTSCAXVersion) then
    objLogFile.WriteLine("INFO: C:\Windows\System32\mstscax.dll is not the desired version")
    ' Check if backup already exists, skip if true
    if not (objFSO.FileExists("C:\TEMP\32\mstscax.dll")) then
      objLogFile.WriteLine("INFO: Backing up C:\Windows\System32\mstscax.dll before replacing.")
      objFSO.CopyFile "C:\Windows\System32\mstscax.dll", "C:\TEMP\32\mstscax.dll", false
    else
      objLogFile.WriteLine("INFO: C:\Windows\System32\mstscax.dll backup already exists.")
    end if
    ' take ownership of file in question
    objShell.Run("takeown /f C:\Windows\System32\mstscax.dll"),0,true
    objShell.Run("icacls C:\Windows\System32\mstscax.dll /grant administrators:F"),0,true
    ' replace file
    objLogFile.WriteLine("INFO: Copying desired version of C:\Windows\System32\mstscax.dll.")
    objFSO.CopyFile DesiredFilePath + "\32\mstscax.dll", "C:\Windows\System32\mstscax.dll", true
  else
    objLogFile.WriteLine("INFO: C:\Windows\System32\mstscax.dll is the desired version")
  end if

  ' Now for SysWOW64
  if (CurrentMSTSCVersion64 <> DesiredMSTSCVersion64) then
    objLogFile.WriteLine("INFO: C:\Windows\SysWOW64\mstsc.exe is not the desired version")

    ' Check if backup already exists, skip if true
    if not (objFSO.FileExists("C:\TEMP\64\mstsc.exe")) then
      objLogFile.WriteLine("INFO: Backing up C:\Windows\SysWOW64\mstsc.exe before replacing.")
      objFSO.CopyFile "C:\Windows\SysWOW64\mstsc.exe", "C:\TEMP\64\mstsc.exe", false
    else
      objLogFile.WriteLine("INFO: C:\Windows\SysWOW64\mstsc.exe backup already exists.")
    end if
    ' take ownership of file in question
    objShell.Run("takeown /f C:\Windows\SysWOW64\mstsc.exe"),0,true
    objShell.Run("icacls C:\Windows\SysWOW64\mstsc.exe /grant administrators:F"),0,true
    ' replace file
    objLogFile.WriteLine("INFO: Copying desired version of C:\Windows\SysWOW64\mstsc.exe.")
    objFSO.CopyFile DesiredFilePath + "\64\mstsc.exe", "C:\Windows\SysWOW64\mstsc.exe", true
  else
    objLogFile.WriteLine("INFO: C:\Windows\SysWOW64\mstsc.exe is the desired version")
  end if

  if (CurrentMSTSCAXVersion64 <> DesiredMSTSCAXVersion64) then
    objLogFile.WriteLine("INFO: C:\Windows\SysWOW64\mstscax.dll is not the desired version")
    ' Check if backup already exists, skip if true
    if not (objFSO.FileExists("C:\TEMP\64\mstscax.dll")) then
      objLogFile.WriteLine("INFO: Backing up C:\Windows\System32\mstscax.dll before replacing.")
      objFSO.CopyFile "C:\Windows\SysWOW64\mstscax.dll", "C:\TEMP\64\mstscax.dll", false
    else
      objLogFile.Writeline("INFO: C:\Windows\SysWOW64\mstscax.dll backup already exists.")
    end if
    ' take ownership of file in question
    objShell.Run("takeown /f C:\Windows\SysWOW64\mstscax.dll"),0,true
    objShell.Run("icacls C:\Windows\SysWOW64\mstscax.dll /grant administrators:F"),0,true
    ' replace file
    objLogFile.WriteLine("INFO: Copying desired version of C:\Windows\SysWOW64\mstscax.dll.")
    objFSO.CopyFile DesiredFilePath + "\64\mstscax.dll", "C:\Windows\SysWOW64\mstscax.dll", true
  else
    objLogFile.WriteLine("INFO: C:\Windows\SysWOW64\mstscax.dll is the desired version")
  end if
else
  if not (objFSO.FolderExists(DesiredFilePath)) then
    objLogFile.WriteLine("ERROR: Directory " + DesiredFilePath + " not found.")
  end if
  if not (objFSO.FileExists(DesiredFilePath + "\32\mstsc.exe")) then
    objLogFile.WriteLine("ERROR: File " + DesiredFilePath + "\32\mstsc.exe not found.")
  end if
  if not (objFSO.FileExists(DesiredFilePath + "\32\mstscax.dll")) then
    objLogFile.WriteLine("ERROR: File " + DesiredFilePath + "\32\mstscax.dll not found.")
  end if
  if not (objFSO.FileExists(DesiredFilePath + "\64\mstsc.exe")) then
    objLogFile.WriteLine("ERROR: File " + DesiredFilePath + "\64\mstsc.exe not found.")
  end if
  if not (objFSO.FileExists(DesiredFilePath + "\64\mstscax.dll")) then
    objLogFile.WriteLine("ERROR: File " + DesiredFilePath + "\64\mstscax.dll not found.")
  end if
end if
