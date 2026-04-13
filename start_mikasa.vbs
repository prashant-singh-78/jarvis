Set shell = CreateObject("WScript.Shell")
Set svc = GetObject("winmgmts:\\.\\root\\cimv2")

scriptPath = "C:\\Users\\prash\\OneDrive\\Desktop\\mikasa.jarvise\\voice_tabs.ps1"

query = "Select * from Win32_Process Where Name='powershell.exe' and CommandLine like '%voice_tabs.ps1%'"
Set processes = svc.ExecQuery(query)
If processes.Count > 0 Then
    WScript.Quit 0
End If

shell.CurrentDirectory = "C:\\Users\\prash\\OneDrive\\Desktop\\mikasa.jarvise"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" & scriptPath & "\""
shell.Run cmd, 0, False
