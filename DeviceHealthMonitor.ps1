# =====================================================
# Device Health Monitor (Production Version)
# =====================================================

# -------- CONFIG --------
$BasePath  = "C:\DeviceHealth"
$ScriptUrl = "https://raw.githubusercontent.com/it-garrod/device-health-monitor/refs/heads/main/DeviceHealthMonitor.ps1"

$TenantId = "25893136-718d-41e2-9817-8cd07ba8a0fc"
$ClientId = "6cc9d032-be18-4d2c-ad99-1090735868a3"

$Recipient = "it@garrod.ph"
$Sender    = "laptops@garrod.ph"

$ScriptVersion = "1.0.0"

# -------- PATHS --------
$LocalScript = "$BasePath\DeviceHealthMonitor.ps1"
$TempScript  = "$BasePath\script_new.ps1"

$Computer = $env:COMPUTERNAME
$Date     = Get-Date -Format "yyyy-MM-dd"
$Report   = "$BasePath\DeviceHealth_$Computer_$Date.txt"

# -------- ENSURE DIRECTORY --------
New-Item -ItemType Directory -Path $BasePath -Force | Out-Null

# -------- SELF UPDATE --------
try {
    Invoke-WebRequest -Uri $ScriptUrl -OutFile $TempScript -UseBasicParsing

    if ((Get-Item $TempScript).Length -gt 1000) {
        Copy-Item $TempScript $LocalScript -Force
    }
}
catch {
    Write-Output "Update failed — continuing with existing script"
}

# -------- BUILD REPORT --------
"Device Health Report" | Out-File $Report
"Computer : $Computer" | Out-File $Report -Append
"Date     : $Date" | Out-File $Report -Append
"Version  : $ScriptVersion" | Out-File $Report -Append
"===============================" | Out-File $Report -Append
"" | Out-File $Report -Append

# -------- DISK HEALTH --------
"Disk Health (SMART):" | Out-File $Report -Append

$Smart = Get-WmiObject -Namespace root\wmi `
    -Class MSStorageDriver_FailurePredictStatus `
    -ErrorAction SilentlyContinue

if ($Smart) {
    foreach ($d in $Smart) {
        if ($d.PredictFailure) {
            "[WARN] Disk failure predicted" | Out-File $Report -Append
        } else {
            "[OK] Disk OK" | Out-File $Report -Append
        }
    }
}
else {
    "[INFO] SMART data unavailable" | Out-File $Report -Append
}

"" | Out-File $Report -Append

# -------- DISK USAGE --------
"Disk Usage:" | Out-File $Report -Append

Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {

    $Free = [Math]::Round($_.FreeSpace / 1GB, 2)
    $Size = [Math]::Round($_.Size / 1GB, 2)

    "Drive $($_.DeviceID): $Free GB free of $Size GB" |
        Out-File $Report -Append

    if ($Free -lt 10) {
        "[WARN] Low disk space on $($_.DeviceID)" |
            Out-File $Report -Append
    }
}

"" | Out-File $Report -Append

# -------- BATTERY STATUS --------
"Battery Status:" | Out-File $Report -Append

$Battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue

if (-not $Battery) {
    "[INFO] No battery detected" | Out-File $Report -Append
}
else {
    foreach ($b in $Battery) {

        $Charge = $b.EstimatedChargeRemaining
        $Status = $b.BatteryStatus

        "Charge: $Charge%" | Out-File $Report -Append

        $PowerState = switch ($Status) {
            1 { "Discharging" }
            2 { "Connected (AC)" }
            3 { "Fully Charged" }
            6 { "Charging" }
            Default { "Unknown" }
        }

        "State : $PowerState" | Out-File $Report -Append

        if ($Charge -le 20) {
            "[WARN] Battery below 20%" | Out-File $Report -Append
        }

        "" | Out-File $Report -Append
    }
}

# -------- SYSTEM UPTIME --------
"System Uptime:" | Out-File $Report -Append

$LastBoot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
$Uptime   = (Get-Date) - $LastBoot

"Uptime: $($Uptime.Days) days" | Out-File $Report -Append

if ($Uptime.Days -ge 14) {
    "[WARN] System running >14 days (restart recommended)" |
        Out-File $Report -Append
}

"" | Out-File $Report -Append

# -------- PENDING REBOOT --------
"Pending Reboot:" | Out-File $Report -Append

$RebootPending = $false

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
    $RebootPending = $true
}

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
    $RebootPending = $true
}

if ($RebootPending) {
    "[WARN] Pending reboot detected" | Out-File $Report -Append
}
else {
    "[OK] No pending reboot" | Out-File $Report -Append
}

# -------- ISSUE DETECTION --------
$HasIssue = Select-String -Path $Report `
    -Pattern "\[WARN\]|\[CRITICAL\]|FAILURE|unavailable" `
    -Quiet

if (-not $HasIssue) {
    Write-Output "No issues detected — skipping email"
    
    "[$(Get-Date)] Healthy - no alert sent" |
        Out-File "$BasePath\run.log" -Append
    
    return
}

# -------- LOAD SECRET --------
$Cred = Import-Clixml "$BasePath\graph_cred.xml"
$ClientSecret = $Cred.GetNetworkCredential().Password

# -------- GRAPH AUTH --------
$TokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

$TokenBody = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}

$Token = Invoke-RestMethod -Method POST -Uri $TokenUri -Body $TokenBody
$AccessToken = $Token.access_token

# -------- EMAIL --------
$ReportText = Get-Content $Report -Raw

$Attachment = [Convert]::ToBase64String(
    [Text.Encoding]::UTF8.GetBytes($ReportText)
)

$MailBody = @{
    message = @{
        subject = "[ATTENTION] Device Health Issue - $Computer"
        body = @{
            contentType = "Text"
            content = "Issues detected on device. See attached report."
        }
        toRecipients = @(
            @{ emailAddress = @{ address = $Recipient } }
        )
        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name = "DeviceHealth_$Computer.txt"
                contentBytes = $Attachment
                contentType = "text/plain"
            }
        )
    }
    saveToSentItems = $false
}

try {
    Invoke-RestMethod `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" `
        -Headers @{
            Authorization = "Bearer $AccessToken"
            "Content-Type" = "application/json"
        } `
        -Body ($MailBody | ConvertTo-Json -Depth 10)

    "[$(Get-Date)] ALERT sent" |
        Out-File "$BasePath\run.log" -Append
}
catch {
    "[$(Get-Date)] ERROR sending email" |
        Out-File "$BasePath\run.log" -Append
}
