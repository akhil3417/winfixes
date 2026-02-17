# ===============================
# CONFIG
# ===============================
$Username  = "shareuser"
$ShareName = "UserShare"
$BasePath  = "C:\Users\$Username\Shared"

# ===============================
# CREATE USER (SECURE PROMPT)
# ===============================
if (-not (Get-LocalUser -Name $Username -ErrorAction SilentlyContinue)) {
    $SecurePass = Read-Host "Enter password for $Username" -AsSecureString
    New-LocalUser -Name $Username -Password $SecurePass -PasswordNeverExpires
    Add-LocalGroupMember -Group "Users" -Member $Username
}

# ===============================
# CREATE FOLDER
# ===============================
if (-not (Test-Path $BasePath)) {
    New-Item -ItemType Directory -Path $BasePath | Out-Null
}

# ===============================
# NTFS PERMISSIONS (CLEAN)
# ===============================
icacls $BasePath /inheritance:e | Out-Null
icacls $BasePath /grant "$Username:(OI)(CI)M" | Out-Null

# ===============================
# SMB HARDENING
# ===============================
Set-SmbServerConfiguration 
    -EnableSMB1Protocol $false 
    -RejectUnencryptedAccess $true 
    -EnableSecuritySignature $true 
    -ForceSecuritySignature $true 
    -Confirm:$false

Set-SmbClientConfiguration 
    -EnableInsecureGuestLogons $false 
    -Confirm:$false

# ===============================
# FIREWALL RULES
# ===============================
Get-NetFirewallRule -DisplayGroup "Network Discovery" | Enable-NetFirewallRule
Get-NetFirewallRule -DisplayGroup "File And Printer Sharing" | Enable-NetFirewallRule

# ===============================
# CREATE SHARE
# ===============================
if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
    New-SmbShare 
        -Name $ShareName 
        -Path $BasePath 
        -FullAccess $Username `
        -FolderEnumerationMode AccessBased
}

# ===============================
# DONE
# ===============================
Write-Host ""
Write-Host "✔ SMB Share Ready"
Write-Host "Share Path : \\$env:COMPUTERNAME\$ShareName"
Write-Host "Login As  : $env:COMPUTERNAME\$Username"
