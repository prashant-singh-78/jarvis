# Ensures Mikasa voice tabs launch immediately after logon (no delay).
# Attempts a per-user scheduled task; if that fails (e.g., access denied), falls back to HKCU Run key.

$taskName = "Mikasa Voice Tabs"
$target = "C:\\Users\\prash\\OneDrive\\Desktop\\mikasa.jarvise\\start_mikasa.vbs"
$runKeyPath = "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"
$runValueName = "MikasaVoiceTabs"
$runCommand = "wscript.exe `"$target`""

if (-not (Test-Path $target)) {
    Write-Error "start_mikasa.vbs not found at $target. Run this script from the project folder."
    exit 1
}

function Register-LogonTask {
    $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$target`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Launch Mikasa voice tabs right at user logon" -ErrorAction Stop | Out-Null
}

function Register-RunKey {
    if (-not (Test-Path $runKeyPath)) { New-Item -Path $runKeyPath -Force | Out-Null }
    New-ItemProperty -Path $runKeyPath -Name $runValueName -Value $runCommand -PropertyType String -Force | Out-Null
}

try {
    Register-LogonTask
    Write-Host "Scheduled task '$taskName' set to run immediately at logon." -ForegroundColor Green
}
catch {
    Write-Warning "Scheduled task failed (`$($_.Exception.Message)`), falling back to HKCU Run key."
    try {
        Register-RunKey
        Write-Host "Run key '$runValueName' added under HKCU\\...\\Run for immediate logon launch." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to set any auto-start mechanism: $_"
        exit 1
    }
}
