Restart-Service LanmanServer -Force
    Restart-Service LanmanWorkstation -Force

    Show-Info -Mode "SECURE" -Password $PlainPassword
}

# ---------------- GUEST MODE ----------------
function Guest-SMB {
    Write-Host "n[ GUEST MODE – PASSWORDLESS ]n"

    Create-Folder

    icacls $SharePath /inheritance:e | Out-Null
    icacls $SharePath /grant "Everyone:(OI)(CI)M" | Out-Null

    reg add "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" 
        /v everyoneincludesanonymous /t REG_DWORD /d 1 /f | Out-Null

    reg add "HKLM\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" 
        /v AllowInsecureGuestAuth /t REG_DWORD /d 1 /f | Out-Null

    Ensure-Services
    Enable-Firewall
    Enable-NetBIOS

    if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $ShareName -Path $SharePath -FullAccess Everyone
    }

    Restart-Service LanmanServer -Force
    Restart-Service LanmanWorkstation -Force

    Show-Info -Mode "GUEST"
}

# ---------------- RESET CLIENT ----------------
function Reset-Client {
    Write-Host "n[ RESETTING CLIENT SMB CREDENTIALS ]n"

    net use * /delete /y | Out-Null

    cmdkey /list | ForEach-Object {
        if ($_ -match "Target") {
            $t = ($_ -split ":")[1].Trim()
            cmdkey /delete:$t | Out-Null
        }
    }

    Write-Host "✔ All SMB credentials cleared"
}

# ---------------- MENU ----------------
Write-Host ""
Write-Host "=========== SMB MANAGEMENT MENU ==========="
Write-Host "1. Secure Share (Password, No Prompts Later)"
Write-Host "2. Guest Share (Passwordless – Insecure)"
Write-Host "3. Reset Client Credentials"
Write-Host "4. Exit"
Write-Host "==========================================="

$Choice = Read-Host "Choose option (1-4)"

switch ($Choice) {
    "1" { Secure-SMB }
    "2" { Guest-SMB }
    "3" { Reset-Client }
    default { Write-Host "Exit." }
}

Stop-Transcript
Write-Host "Log saved to $LogFile"
