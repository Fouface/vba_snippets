Attribute VB_Name = "z_shell_sendkeys"
Sub UsingSendKeysWithWait()

Call Shell("C:\Windows\system32\Notepad.Exe", vbNormalFocus)
Application.Wait (Now() + TimeValue("00:00:10"))
Call SendKeys("This is Some Text", True)

End Sub




