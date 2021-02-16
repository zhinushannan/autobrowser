On Error Resume Next

Set Wshshell=WScript.CreateObject("WScript.Shell")

WshShell.run"""C:\Program Files (x86)\Tencent\QQ\Bin\QQScLauncher.exe"" /uin:3528275423 /quicklunch:6886EE5EFC1860EC08125DA0051541B7CEA41EE7D9B5D6C4FFDEE949B8CB5B01A0FB09115D98A1A9"

WScript.Sleep 1000

WshShell.SendKeys"^v"

WScript.Sleep 1000

WshShell.SendKeys "%s"

WScript.Sleep 1000
