Dim WshShell, paused
Set WshShell = CreateObject("WScript.Shell")
paused = False

Do
    If Not paused Then
        WshShell.SendKeys "attack the d point"
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 10 ' Adjust the sleep time as needed
    End If
    
    If WshShell.AppActivate("Command Prompt") Then
        If WshShell.SendKeys("^p") Then ' Ctrl + P to pause
            paused = True
        ElseIf WshShell.SendKeys("^s") Then ' Ctrl + S to restart
            paused = False
        End If
    End If
Loop
