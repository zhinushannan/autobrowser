On Error Resume Next

Set Wshshell=WScript.CreateObject("WScript.Shell")

WScript.Sleep 1000

WshShell.run"""C:\Program Files (x86)\Tencent\QQ\Bin\QQScLauncher.exe"" /uin:1377875184 /quicklunch:DAE2992EB7205B7043A8611627A0F8D54B17D727B40D3E6463DAD4924D6851C88C012C2BA0ECDF07"

WScript.Sleep 1000

WshShell.SendKeys"^v"

WScript.Sleep 1000

WshShell.SendKeys "%s"

WScript.Sleep 1000
