param(
    [string]$TranscriptPath
)

$ErrorActionPreference = 'Stop'

if ($TranscriptPath) {
    try {
        Start-Transcript -Path $TranscriptPath -Force | Out-Null
    }
    catch {
        Write-Warning "Failed to start transcript. $($_.Exception.Message)"
    }
}

try {
    $method = if ($env:IMPACTIQ_NOTIFY_METHOD) { $env:IMPACTIQ_NOTIFY_METHOD.Trim().ToUpperInvariant() } else { 'SMTP' }

    $to = @()
    if ($env:IMPACTIQ_NOTIFY_TO) {
        $to += ($env:IMPACTIQ_NOTIFY_TO -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if ($env:IMPACTIQ_NOTIFY_SMS_TO) {
        $to += ($env:IMPACTIQ_NOTIFY_SMS_TO -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if (-not $to -or $to.Count -eq 0) {
        throw "No recipients configured. Set IMPACTIQ_NOTIFY_TO and/or IMPACTIQ_NOTIFY_SMS_TO."
    }

    $subject = if ($env:IMPACTIQ_NOTIFY_SUBJECT) { $env:IMPACTIQ_NOTIFY_SUBJECT } else { '[ImpactIQ] Notification test' }
    $body = "ImpactIQ notification test succeeded.`nHost: $env:COMPUTERNAME`nTime: $(Get-Date -Format o)"

    if ($method -eq 'OUTLOOK') {
        $outlook = $null
        $mailItem = $null

        try {
            $outlook = New-Object -ComObject Outlook.Application
            $mailItem = $outlook.CreateItem(0)
            $mailItem.Subject = $subject
            $mailItem.Body = $body
            $mailItem.To = ($to -join ';')
            $mailItem.Send()
        }
        finally {
            if ($mailItem) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailItem) }
            if ($outlook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
        }
    }
    elseif ($method -eq 'SMTP') {
        if (-not $env:IMPACTIQ_NOTIFY_SMTP_SERVER -or -not $env:IMPACTIQ_NOTIFY_FROM) {
            throw "SMTP mode requires IMPACTIQ_NOTIFY_SMTP_SERVER and IMPACTIQ_NOTIFY_FROM."
        }

        $smtpPort = 587
        if ($env:IMPACTIQ_NOTIFY_SMTP_PORT -and ($env:IMPACTIQ_NOTIFY_SMTP_PORT -as [int])) {
            $smtpPort = [int]$env:IMPACTIQ_NOTIFY_SMTP_PORT
        }

        $useSsl = $true
        if ($env:IMPACTIQ_NOTIFY_SMTP_USE_SSL -match '^(?i:false|0|no)$') {
            $useSsl = $false
        }

        $mailParams = @{
            SmtpServer = $env:IMPACTIQ_NOTIFY_SMTP_SERVER
            Port       = $smtpPort
            UseSsl     = $useSsl
            From       = $env:IMPACTIQ_NOTIFY_FROM
            To         = ($to -join ',')
            Subject    = $subject
            Body       = $body
            ErrorAction= 'Stop'
        }

        if ($env:IMPACTIQ_NOTIFY_SMTP_USERNAME -and $env:IMPACTIQ_NOTIFY_SMTP_PASSWORD) {
            $securePassword = ConvertTo-SecureString $env:IMPACTIQ_NOTIFY_SMTP_PASSWORD -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($env:IMPACTIQ_NOTIFY_SMTP_USERNAME, $securePassword)
            $mailParams.Credential = $credential
        }

        Send-MailMessage @mailParams
    }
    else {
        throw "Unknown IMPACTIQ_NOTIFY_METHOD '$method'. Use SMTP or OUTLOOK."
    }

    Write-Host "[INFO] Notification verification sent via $method to: $($to -join ', ')"
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {
    }
}
