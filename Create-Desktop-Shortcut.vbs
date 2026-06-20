' ============================================================
'  Creates a "SOP Planning" desktop shortcut.
'  Double-click this file ONCE to create the shortcut.
'  After that, use the desktop icon to start the tool.
' ============================================================

Set oWS = WScript.CreateObject("WScript.Shell")
strDesktop = oWS.SpecialFolders("Desktop")
strProjectFolder = oWS.CurrentDirectory

Set oLink = oWS.CreateShortcut(strDesktop & "\SOP Planning.lnk")
oLink.TargetPath = strProjectFolder & "\Start-SOP-Planning.bat"
oLink.WorkingDirectory = strProjectFolder
oLink.WindowStyle = 1
oLink.Description = "Start the SOP Planning Excel add-in"

' Uses Excel's own icon so it's instantly recognisable.
' Replace this path with a custom .ico file if you have one,
' e.g. oLink.IconLocation = strProjectFolder & "\icon.ico"
oLink.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE,0"

oLink.Save

MsgBox "Shortcut created on your Desktop: 'SOP Planning'." & vbCrLf & vbCrLf & _
       "Double-click it any time to start the tool.", vbInformation, "SOP Planning Setup"