#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Digest Email Campaign - Basic Authentication
.DESCRIPTION
    Bulk email sending with Basic SMTP authentication featuring:
    - Secure credential handling (multiple options)
    - CSV data import with validation
    - Template-based HTML emails
    - Inline images and PDF attachments
    - Batched sending with rate limiting
    - Checkpoint-based resume capability
.PARAMETER ConfigMode
    'Test' for reduced batch size, 'Production' for full campaign
.PARAMETER CredentialMethod
    How to obtain SMTP credentials:
    - 'Prompt'      : Interactive prompt (default, most secure for manual runs)
    - 'Environment' : From environment variables
    - 'CredentialManager' : From Windows Credential Manager (requires CredentialManager module)
    - 'SecureFile'  : From encrypted XML file (user-specific DPAPI encryption)
.PARAMETER SmtpUsername
    SMTP username (required for Environment method, optional override for others)
.NOTES
    Requires: M365DigestEmailModule.psm1

    SECURITY: This script does NOT store credentials in source code.
    Choose a credential method appropriate for your use case:
    - Interactive use: 'Prompt' (default)
    - Scheduled tasks: 'SecureFile' or 'CredentialManager'
    - CI/CD pipelines: 'Environment'
.AUTHOR
    Jan Huebener
.VERSION
    1.1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Test', 'Production')]
    [string]$ConfigMode = 'Production',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Prompt', 'Environment', 'CredentialManager', 'SecureFile')]
    [string]$CredentialMethod = 'Prompt',

    [Parameter(Mandatory = $false)]
    [string]$SmtpUsername,

    [Parameter(Mandatory = $false)]
    [string]$SecureFilePath = "$env:USERPROFILE\.m365digest\smtp_credential.xml",

    [Parameter(Mandatory = $false)]
    [string]$CredentialTarget = "M365-Digest-SMTP"
)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Import module
$modulePath = Join-Path $PSScriptRoot "M365DigestEmailModule.psm1"
Import-Module $modulePath -Force -Verbose

# ============================================================================
# PATH CONFIGURATION - UPDATE THESE FOR YOUR ENVIRONMENT
# ============================================================================

$csvPath = "C:\Temp\recipients.csv"                              # Recipient CSV file
$htmlTemplate = Join-Path $PSScriptRoot "M365_Digest_Template.htm"  # HTML template
$checkpointFile = "C:\Temp\smtp_send_checkpoint_m365digest.txt"   # Checkpoint for resume
$failedRecipientsFile = "C:\Temp\failed_recipients.csv"           # Failed emails log

# Inline Images (CID must match template references)
$inlineImages = @(
    @{
        ContentId = 'datagroup_logo'
        FilePath  = 'C:\temp\datagroup_logo.png'
    },
    @{
        ContentId = 'm365_icon'
        FilePath  = Join-Path $PSScriptRoot 'm365_icon.png'
    },
    @{
        ContentId = 'exchange_icon'
        FilePath  = Join-Path $PSScriptRoot 'exchange_icon.png'
    },
    @{
        ContentId = 'sharepoint_icon'
        FilePath  = Join-Path $PSScriptRoot 'sharepoint_icon.png'
    }
)

# Attachments (optional - comment out if not needed)
$attachments = @(
    # "C:\temp\Document1.pdf",
    # "C:\temp\Document2.pdf"
)

# ============================================================================
# SMTP CONFIGURATION - UPDATE THESE FOR YOUR ENVIRONMENT
# ============================================================================

$smtpConfig = @{
    Server    = "smtp.office365.com"
    Port      = 587
    EnableSsl = $true
    From      = "noreply@yourdomain.com"    # UPDATE: Your sender address
    Bcc       = "admin@yourdomain.com"       # UPDATE: BCC for monitoring (optional)
    Subject   = "Microsoft 365 Monthly Digest - What's new?"
}

# Batch Configuration
$batchConfig = @{
    BatchSize     = 20
    WindowMinutes = 3.0
    MaxRetries    = 3
}

# Test Mode - reduced settings for testing
if ($ConfigMode -eq 'Test') {
    Write-Host "`n[TEST MODE] Using reduced batch settings" -ForegroundColor Yellow
    $batchConfig.BatchSize = 2
    $batchConfig.WindowMinutes = 0.1
}

# ============================================================================
# SECURE CREDENTIAL FUNCTIONS
# ============================================================================

function Get-SmtpCredentialSecure {
    <#
    .SYNOPSIS
        Retrieves SMTP credentials using the specified method
    #>
    param(
        [string]$Method,
        [string]$Username,
        [string]$SecureFilePath,
        [string]$CredentialTarget
    )

    switch ($Method) {
        'Prompt' {
            Write-Host "  [Prompt] Enter SMTP credentials..." -ForegroundColor Cyan
            $cred = Get-Credential -Message "Enter SMTP credentials for M365 Digest"
            if (-not $cred) {
                throw "Credential prompt was cancelled"
            }
            return $cred
        }

        'Environment' {
            Write-Host "  [Environment] Loading from environment variables..." -ForegroundColor Cyan
            $envUser = $env:M365_SMTP_USERNAME
            $envPass = $env:M365_SMTP_PASSWORD

            if ([string]::IsNullOrWhiteSpace($envUser)) {
                throw "Environment variable M365_SMTP_USERNAME is not set"
            }
            if ([string]::IsNullOrWhiteSpace($envPass)) {
                throw "Environment variable M365_SMTP_PASSWORD is not set"
            }

            $securePass = ConvertTo-SecureString $envPass -AsPlainText -Force
            return New-Object System.Management.Automation.PSCredential($envUser, $securePass)
        }

        'CredentialManager' {
            Write-Host "  [CredentialManager] Loading from Windows Credential Manager..." -ForegroundColor Cyan

            # Check if CredentialManager module is available
            if (-not (Get-Module -ListAvailable -Name CredentialManager)) {
                throw "CredentialManager module not installed. Install with: Install-Module CredentialManager"
            }

            Import-Module CredentialManager -ErrorAction Stop
            $cred = Get-StoredCredential -Target $CredentialTarget

            if (-not $cred) {
                throw "Credential '$CredentialTarget' not found in Windows Credential Manager. Create it with: New-StoredCredential -Target '$CredentialTarget' -UserName 'user@domain.com' -Password 'password' -Persist LocalMachine"
            }
            return $cred
        }

        'SecureFile' {
            Write-Host "  [SecureFile] Loading from encrypted file..." -ForegroundColor Cyan

            if (-not (Test-Path $SecureFilePath)) {
                Write-Host "`n  Secure credential file not found. Creating one now..." -ForegroundColor Yellow
                Write-Host "  File: $SecureFilePath" -ForegroundColor Gray

                # Ensure directory exists
                $secureDir = Split-Path $SecureFilePath -Parent
                if (-not (Test-Path $secureDir)) {
                    New-Item -Path $secureDir -ItemType Directory -Force | Out-Null
                }

                # Prompt for credentials and save encrypted
                $newCred = Get-Credential -Message "Enter SMTP credentials to save (encrypted with your Windows account)"
                if (-not $newCred) {
                    throw "Credential prompt was cancelled"
                }

                $newCred | Export-Clixml -Path $SecureFilePath
                Write-Host "  [OK] Credentials saved to: $SecureFilePath" -ForegroundColor Green
                Write-Host "  Note: This file can only be decrypted by your Windows account on this machine.`n" -ForegroundColor Gray

                return $newCred
            }

            # Load existing encrypted credential
            $cred = Import-Clixml -Path $SecureFilePath
            return $cred
        }

        default {
            throw "Unknown credential method: $Method"
        }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

try {
    Write-Host "`n" -NoNewline
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "  M365 Monthly Digest Email Campaign (Basic Auth)       " -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "  Mode: $ConfigMode" -ForegroundColor $(if ($ConfigMode -eq 'Test') { 'Yellow' } else { 'Green' })
    Write-Host "  Credential Method: $CredentialMethod" -ForegroundColor Gray
    Write-Host "========================================================`n" -ForegroundColor Cyan

    # ========================================================================
    # STEP 1: Validate file paths
    # ========================================================================
    Write-Host "[1/5] Validating file paths..." -ForegroundColor Cyan

    if (-not (Test-Path $csvPath)) {
        throw "CSV file not found: $csvPath"
    }
    Write-Host "  [OK] CSV file: $csvPath" -ForegroundColor Green

    if (-not (Test-Path $htmlTemplate)) {
        throw "HTML template not found: $htmlTemplate"
    }
    Write-Host "  [OK] HTML template: $htmlTemplate" -ForegroundColor Green

    # ========================================================================
    # STEP 2: Validate inline images
    # ========================================================================
    Write-Host "`n[2/5] Validating inline images..." -ForegroundColor Cyan
    $missingImages = 0
    foreach ($img in $inlineImages) {
        if (Test-Path $img.FilePath) {
            $size = [math]::Round((Get-Item $img.FilePath).Length / 1KB, 1)
            Write-Host "  [OK] $($img.ContentId): ${size}KB" -ForegroundColor Green
        }
        else {
            Write-Host "  [WARN] Missing: $($img.FilePath)" -ForegroundColor Yellow
            $missingImages++
        }
    }
    if ($missingImages -gt 0) {
        Write-Host "  [!] $missingImages image(s) missing - emails may display incorrectly" -ForegroundColor Yellow
    }

    # ========================================================================
    # STEP 3: Get credentials securely
    # ========================================================================
    Write-Host "`n[3/5] Obtaining SMTP credentials..." -ForegroundColor Cyan

    $credential = Get-SmtpCredentialSecure `
        -Method $CredentialMethod `
        -Username $SmtpUsername `
        -SecureFilePath $SecureFilePath `
        -CredentialTarget $CredentialTarget

    # Update SMTP config with credential
    $smtpConfig.Credential = Get-EmailAuthenticationCredential `
        -AuthMethod 'Basic' `
        -Username $credential.UserName `
        -Password $credential.GetNetworkCredential().Password

    # Update From address to match credential username if not explicitly set
    if ($smtpConfig.From -match 'yourdomain.com') {
        $smtpConfig.From = $credential.UserName
        Write-Host "  [INFO] Using credential username as From address" -ForegroundColor Gray
    }

    Write-Host "  [OK] Credentials loaded for: $($credential.UserName)" -ForegroundColor Green

    # ========================================================================
    # STEP 4: Load and validate recipients
    # ========================================================================
    Write-Host "`n[4/5] Loading recipient data..." -ForegroundColor Cyan
    $csvData = Import-Csv -LiteralPath $csvPath -Delimiter ';' -Encoding UTF8

    # Build recipient objects with validation
    $recipients = @()
    $invalidEmails = @()
    $emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

    foreach ($row in $csvData) {
        $email = ($row.email).Trim()

        # Skip empty emails
        if ([string]::IsNullOrWhiteSpace($email)) { continue }

        # Validate email format
        if ($email -notmatch $emailRegex) {
            $invalidEmails += $email
            Write-Host "  [WARN] Invalid email skipped: $email" -ForegroundColor Yellow
            continue
        }

        # Build replacement hashtable for this recipient
        $replacements = @{
            'CARD1_TITLE'   = "New Teams Features"
            'CARD1_CONTENT' = "Microsoft Teams introduces new collaboration features including enhanced meeting recordings and AI-powered meeting summaries."
            'CARD1_LINK'    = "https://admin.microsoft.com/?ref=MessageCenter/:/messages/MC1069560"

            'CARD2_TITLE'   = "Exchange Online Updates"
            'CARD2_CONTENT' = "Enhanced security features now available for Exchange Online mailboxes, including improved phishing protection."
            'CARD2_LINK'    = "https://admin.microsoft.com/?ref=MessageCenter/:/messages/MC1134178"

            'CARD3_TITLE'   = "SharePoint Improvements"
            'CARD3_CONTENT' = "New document management capabilities in SharePoint Online with AI-powered search and classification."
            'CARD3_LINK'    = "https://admin.microsoft.com/?ref=MessageCenter/:/messages/MC1069560"

            'UNSUBSCRIBE_LINK' = "https://www.yourdomain.com/unsubscribe?email=$email"
        }

        # Personalize if DisplayName is available
        if ($row.PSObject.Properties.Name -contains 'DisplayName_email' -and $row.DisplayName_email) {
            $replacements['CARD1_CONTENT'] = "Hello $($row.DisplayName_email), " + $replacements['CARD1_CONTENT']
        }

        $recipients += [PSCustomObject]@{
            Email        = $email
            Replacements = $replacements
        }
    }

    Write-Host "  [OK] Valid recipients: $($recipients.Count)" -ForegroundColor Green
    if ($invalidEmails.Count -gt 0) {
        Write-Host "  [WARN] Invalid emails skipped: $($invalidEmails.Count)" -ForegroundColor Yellow
    }

    if ($recipients.Count -eq 0) {
        throw "No valid recipients found in CSV file"
    }

    # ========================================================================
    # STEP 5: Display summary and confirm
    # ========================================================================
    Write-Host "`n[5/5] Campaign Summary" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "  Recipients:      $($recipients.Count)"
    Write-Host "  Batch Size:      $($batchConfig.BatchSize)"
    Write-Host "  Window Interval: $($batchConfig.WindowMinutes) minutes"
    Write-Host "  Max Retries:     $($batchConfig.MaxRetries)"
    Write-Host "  Sender:          $($smtpConfig.From)"
    Write-Host "  Subject:         $($smtpConfig.Subject)"
    Write-Host "  Checkpoint:      $checkpointFile"
    Write-Host "========================================================`n" -ForegroundColor Cyan

    if ($ConfigMode -eq 'Production' -and $recipients.Count -gt 10) {
        Write-Host "[!] PRODUCTION MODE - Sending to $($recipients.Count) recipients" -ForegroundColor Yellow
        Write-Host "    Press Ctrl+C within 5 seconds to abort..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }

    # Template Configuration
    $templateConfig = @{
        TemplatePath = $htmlTemplate
        Encoding     = 'UTF8'
        InlineImages = $inlineImages
        Attachments  = $attachments
    }

    # Send bulk emails
    $bulkParams = @{
        Recipients      = $recipients
        TemplateConfig  = $templateConfig
        SmtpConfig      = $smtpConfig
        BatchSize       = $batchConfig.BatchSize
        WindowMinutes   = $batchConfig.WindowMinutes
        MaxRetries      = $batchConfig.MaxRetries
        CheckpointPath  = $checkpointFile
    }

    Send-BulkHtmlEmail @bulkParams

    Write-Host "`n========================================================" -ForegroundColor Green
    Write-Host "  CAMPAIGN COMPLETED SUCCESSFULLY" -ForegroundColor Green
    Write-Host "========================================================`n" -ForegroundColor Green

}
catch {
    Write-Host "`n========================================================" -ForegroundColor Red
    Write-Host "  CAMPAIGN FAILED" -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`n  Stack Trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host "========================================================`n" -ForegroundColor Red
    exit 1
}
