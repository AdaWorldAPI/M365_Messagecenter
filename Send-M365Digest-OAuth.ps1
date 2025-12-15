#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Digest Email Campaign - OAuth2 Authentication (Azure AD App)
.DESCRIPTION
    Example script demonstrating bulk email sending with:
    - OAuth2 Client Credentials Flow (Azure AD App Registration)
    - CSV data import
    - Template-based HTML emails
    - Inline images (logo + 3 product icons)
    - PDF attachments
    - Batched sending with rate limiting
.NOTES
    Requires:
    - M365DigestEmailModule.psm1
    - Azure AD App Registration with Mail.Send permission
    - Admin consent granted for the application

    Setup Steps:
    1. Register app in Azure AD: Azure Portal > App registrations > New
    2. Add API permission: Microsoft Graph > Application > Mail.Send
    3. Grant admin consent
    4. Create client secret and note the value
    5. Update configuration below with TenantId, ClientId, ClientSecret
.AUTHOR
    Jan Huebener
.VERSION
    1.0.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Test', 'Production')]
    [string]$ConfigMode = 'Production'
)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Import module
$modulePath = Join-Path $PSScriptRoot "M365DigestEmailModule.psm1"
Import-Module $modulePath -Force -Verbose

# Paths - UPDATE THESE FOR YOUR ENVIRONMENT
$csvPath = "C:\Temp\master_users_all_merged_with_wave4.csv"
$htmlTemplate = Join-Path $PSScriptRoot "M365_Digest_Template.htm"
$checkpointFile = "C:\Temp\smtp_send_checkpoint_m365digest_oauth.txt"

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

# Attachments - UPDATE OR REMOVE AS NEEDED
$attachments = @(
    # "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Smartphone.pdf",
    # "C:\temp\Anleitung_Erstanmeldung_Authentifizierung_mit_Telefon.pdf"
)

# SMTP Configuration
$smtpConfig = @{
    Server    = "smtp.office365.com"
    Port      = 587
    EnableSsl = $true
    From      = "noreply@yourdomain.com"          # UPDATE: Sender address
    Bcc       = "admin@yourdomain.com"            # UPDATE: BCC for monitoring
    Subject   = "Microsoft 365 Monthly Digest - What's new?"
}

# ============================================================================
# OAUTH CONFIGURATION - UPDATE THESE VALUES
# ============================================================================
#
# To get these values:
# 1. Go to Azure Portal > Azure Active Directory > App registrations
# 2. Select your app (or create new one)
# 3. Overview page has: Application (client) ID and Directory (tenant) ID
# 4. Certificates & secrets > New client secret > Copy the VALUE (not ID)
#
# SECURITY WARNING:
# Do NOT commit real secrets to source control!
# Use secure storage methods in production:
# - Azure Key Vault
# - Windows Credential Manager
# - Environment variables
# - Encrypted config files

$oauthConfig = @{
    TenantId     = "your-tenant-id-here"          # Directory (tenant) ID
    ClientId     = "your-client-id-here"          # Application (client) ID
    ClientSecret = "your-client-secret-here"      # Client secret VALUE
    Username     = "noreply@yourdomain.com"       # Must match From address
}

# ============================================================================
# SECURE CREDENTIAL LOADING (PRODUCTION RECOMMENDED)
# ============================================================================
# Uncomment ONE of these sections for production use:

# Option 1: Load from environment variables
# $oauthConfig = @{
#     TenantId     = $env:M365_TENANT_ID
#     ClientId     = $env:M365_CLIENT_ID
#     ClientSecret = $env:M365_CLIENT_SECRET
#     Username     = $env:M365_SENDER_EMAIL
# }

# Option 2: Load from Azure Key Vault (requires Az.KeyVault module)
# Import-Module Az.KeyVault
# $vaultName = "your-keyvault-name"
# $oauthConfig = @{
#     TenantId     = (Get-AzKeyVaultSecret -VaultName $vaultName -Name "M365-TenantId" -AsPlainText)
#     ClientId     = (Get-AzKeyVaultSecret -VaultName $vaultName -Name "M365-ClientId" -AsPlainText)
#     ClientSecret = (Get-AzKeyVaultSecret -VaultName $vaultName -Name "M365-ClientSecret" -AsPlainText)
#     Username     = (Get-AzKeyVaultSecret -VaultName $vaultName -Name "M365-SenderEmail" -AsPlainText)
# }

# Option 3: Load from encrypted file (Windows DPAPI - user-specific)
# $configPath = "C:\Secure\oauth_config.xml"
# if (Test-Path $configPath) {
#     $encrypted = Import-Clixml $configPath
#     $oauthConfig = @{
#         TenantId     = $encrypted.TenantId
#         ClientId     = $encrypted.ClientId
#         ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
#             [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($encrypted.ClientSecret)
#         )
#         Username     = $encrypted.Username
#     }
# }

# ============================================================================
# Batch Configuration
# ============================================================================

$batchConfig = @{
    BatchSize     = 20
    WindowMinutes = 3.0
    MaxRetries    = 3
}

# Test Mode Configuration - reduced batch for testing
if ($ConfigMode -eq 'Test') {
    Write-Host "`n[TEST MODE] Using reduced batch settings" -ForegroundColor Yellow
    $batchConfig.BatchSize = 2
    $batchConfig.WindowMinutes = 0.1
}

# ============================================================================
# VALIDATION
# ============================================================================

function Test-OAuthConfiguration {
    param([hashtable]$Config)

    $valid = $true

    if ($Config.TenantId -match 'your-tenant-id|^$') {
        Write-Host "  [ERROR] TenantId not configured" -ForegroundColor Red
        $valid = $false
    }
    if ($Config.ClientId -match 'your-client-id|^$') {
        Write-Host "  [ERROR] ClientId not configured" -ForegroundColor Red
        $valid = $false
    }
    if ($Config.ClientSecret -match 'your-client-secret|^$') {
        Write-Host "  [ERROR] ClientSecret not configured" -ForegroundColor Red
        $valid = $false
    }
    if ($Config.Username -match 'yourdomain.com|^$') {
        Write-Host "  [ERROR] Username (sender email) not configured" -ForegroundColor Red
        $valid = $false
    }

    return $valid
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

try {
    Write-Host "`n" -NoNewline
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  M365 Monthly Digest Email Campaign (OAuth2)  " -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  Mode: $ConfigMode" -ForegroundColor $(if ($ConfigMode -eq 'Test') { 'Yellow' } else { 'Green' })
    Write-Host "================================================`n" -ForegroundColor Cyan

    # Validate OAuth configuration
    Write-Host "[1/6] Validating OAuth configuration..." -ForegroundColor Cyan
    if (-not (Test-OAuthConfiguration -Config $oauthConfig)) {
        throw "OAuth configuration is incomplete. Please update the script with your Azure AD app credentials."
    }
    Write-Host "  [OK] OAuth configuration valid" -ForegroundColor Green

    # Validate file paths
    Write-Host "`n[2/6] Validating file paths..." -ForegroundColor Cyan

    if (-not (Test-Path $csvPath)) {
        throw "CSV file not found: $csvPath"
    }
    Write-Host "  [OK] CSV file: $csvPath" -ForegroundColor Green

    if (-not (Test-Path $htmlTemplate)) {
        throw "HTML template not found: $htmlTemplate"
    }
    Write-Host "  [OK] HTML template: $htmlTemplate" -ForegroundColor Green

    # Validate inline images
    Write-Host "`n[3/6] Validating inline images..." -ForegroundColor Cyan
    $missingImages = 0
    foreach ($img in $inlineImages) {
        if (Test-Path $img.FilePath) {
            $size = [math]::Round((Get-Item $img.FilePath).Length / 1KB, 1)
            Write-Host "  [OK] $($img.ContentId): $($img.FilePath) (${size}KB)" -ForegroundColor Green
        }
        else {
            Write-Host "  [WARN] Missing: $($img.FilePath)" -ForegroundColor Yellow
            $missingImages++
        }
    }
    if ($missingImages -gt 0) {
        Write-Host "  [!] $missingImages image(s) missing - emails may display incorrectly" -ForegroundColor Yellow
    }

    # Validate attachments
    Write-Host "`n[4/6] Validating attachments..." -ForegroundColor Cyan
    if ($attachments.Count -eq 0) {
        Write-Host "  [INFO] No attachments configured" -ForegroundColor Gray
    }
    else {
        foreach ($attachment in $attachments) {
            if (Test-Path $attachment) {
                $size = [math]::Round((Get-Item $attachment).Length / 1KB, 1)
                Write-Host "  [OK] $attachment (${size}KB)" -ForegroundColor Green
            }
            else {
                Write-Host "  [WARN] Missing: $attachment" -ForegroundColor Yellow
            }
        }
    }

    # Acquire OAuth token
    Write-Host "`n[5/6] Acquiring OAuth2 token..." -ForegroundColor Cyan
    Write-Host "  Tenant: $($oauthConfig.TenantId)" -ForegroundColor Gray
    Write-Host "  Client: $($oauthConfig.ClientId)" -ForegroundColor Gray

    $credential = Get-EmailAuthenticationCredential `
        -AuthMethod 'OAuth' `
        -Username $oauthConfig.Username `
        -TenantId $oauthConfig.TenantId `
        -ClientId $oauthConfig.ClientId `
        -ClientSecret $oauthConfig.ClientSecret

    $smtpConfig.Credential = $credential
    $smtpConfig.From = $oauthConfig.Username
    Write-Host "  [OK] OAuth token acquired successfully" -ForegroundColor Green

    # Load recipients
    Write-Host "`n[6/6] Loading recipient data..." -ForegroundColor Cyan
    $csvData = Import-Csv -LiteralPath $csvPath -Delimiter ';' -Encoding UTF8

    # Build recipient objects with template replacements
    $recipients = @()
    foreach ($row in $csvData) {
        $email = ($row.email).Trim()
        if ([string]::IsNullOrWhiteSpace($email)) { continue }

        # Build replacement hashtable for this recipient
        # UPDATE THESE to match your actual content
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
            $name = $row.DisplayName_email
            $replacements['CARD1_CONTENT'] = "Hello $name, " + $replacements['CARD1_CONTENT']
        }

        $recipients += [PSCustomObject]@{
            Email        = $email
            Replacements = $replacements
        }
    }

    Write-Host "  [OK] Loaded $($recipients.Count) recipients" -ForegroundColor Green

    # Display summary before sending
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "  CAMPAIGN SUMMARY" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  Recipients: $($recipients.Count)"
    Write-Host "  Batch Size: $($batchConfig.BatchSize)"
    Write-Host "  Window Interval: $($batchConfig.WindowMinutes) minutes"
    Write-Host "  Max Retries: $($batchConfig.MaxRetries)"
    Write-Host "  Sender: $($smtpConfig.From)"
    Write-Host "  Subject: $($smtpConfig.Subject)"
    Write-Host "  Auth: OAuth2 (Client Credentials)"
    Write-Host "================================================`n" -ForegroundColor Cyan

    if ($ConfigMode -eq 'Production') {
        Write-Host "[!] PRODUCTION MODE - Sending to ALL recipients" -ForegroundColor Yellow
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

    Write-Host "`n================================================" -ForegroundColor Green
    Write-Host "  CAMPAIGN COMPLETED SUCCESSFULLY" -ForegroundColor Green
    Write-Host "================================================`n" -ForegroundColor Green

}
catch {
    Write-Host "`n================================================" -ForegroundColor Red
    Write-Host "  CAMPAIGN FAILED" -ForegroundColor Red
    Write-Host "================================================" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`n  Stack Trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host "================================================`n" -ForegroundColor Red
    exit 1
}
