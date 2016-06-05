' -----------------------
'  DapoDog
'  Backup and Restore Aplikasi Dapodikmen
'  Copyleft 2016 wagungs@gmail.com
'
'
'Run As Administrator

Set WshShell = WScript.CreateObject("WScript.Shell")
  If WScript.Arguments.length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe", """" & _
  WScript.ScriptFullName & """" &_
  " RunAsAdministrator", , "runas", 1
  Wscript.Quit
  End if

'Mengehentikan Service 
'Stop Service
strServiceName = "DapodikmenDB"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
objService.StopService()
Next
 
 
'Stop Service
strServiceName = "DapodikmenWebSrv"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
objService.StopService()
Next
 
'Copy file menggunakan Robocopy
Set objShell = CreateObject("Wscript.Shell") 
objSource = InputBox("Enter Source")
objDestination = InputBox("Enter Destination")
objCommand = "RoboCopy.Exe " & Chr(34) & objSource & Chr(34) & " " & Chr(34) & objDestination & Chr(34) & " /MIR /SEC /COPYALL /R:0 /W:0"
objShell.Run(objCommand)
MsgBox "Done"

'Memulai Service
'Start Service
strServiceName = "DapodikmenDB"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
objService.StartService()
Next

'Start Service
strServiceName = "DapodikmenWebSrv"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
objService.StartService()
Next