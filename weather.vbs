Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c C:\Users\bushm\Desktop\weather.bat"
oShell.Run strArgs, 0, false
