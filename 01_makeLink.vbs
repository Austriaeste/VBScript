Option Explicit

Dim WSH,sc

Set WSH=CreateObject("WScript.Shell")

Set sc = WSH.CreateShortcut("C:\Users\Austria\Desktop\フォルダ名\notepad.lnk")
sc.TargetPath = "C:\WINDOWS\notepad.exe"
sc.save

Set sc = Nothing
Set WSH = nothing