#Requires -Version 5.1
<#
.SYNOPSIS
    M365 Digest Email Sender - Modular HTML Email System with OAuth/Basic Auth
.DESCRIPTION
    Professional email sending system with:
    - Separate HTML template processing
    - Inline image embedding with AlternateView
    - OAuth2 and Basic Auth support
    - Batched sending with rate limiting
    - Checkpoint-based resume capability
    - Comprehensive error handling and retry logic
    - Email and URL validation
    - Structured JSON logging
    - Failed recipients tracking
.AUTHOR
    Jan Huebener
.VERSION
    1.1.0
#>

# ============================================================================
# MODULE: Logging Infrastructure
# ============================================================================

# Script-level logging configuration
$script:LogConfig = @{
    Enabled       = $true
    LogFilePath   = $null
    LogLevel      = 'INFO'
    JsonFormat    = $true
    CorrelationId = $null
}

function Set-DigestLogConfig {
    <#
    .SYNOPSIS
        Configures logging settings for the module
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$LogFilePath,

        [Parameter(Mandatory = $false)]
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$LogLevel = 'INFO',

        [Parameter(Mandatory = $false)]
        [bool]$JsonFormat = $true,

        [Parameter(Mandatory = $false)]
        [string]$CorrelationId
    )

    $script:LogConfig.LogFilePath = $LogFilePath
    $script:LogConfig.LogLevel = $LogLevel
    $script:LogConfig.JsonFormat = $JsonFormat
    $script:LogConfig.CorrelationId = $CorrelationId ?? [guid]::NewGuid().ToString('N').Substring(0, 8)
}

function Write-DigestLog {
    <#
    .SYNOPSIS
        Writes a structured log entry
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [hashtable]$Properties = @{},

        [Parameter(Mandatory = $false)]
        [string]$Email
    )

    $levelOrder = @{ 'DEBUG' = 0; 'INFO' = 1; 'WARN' = 2; 'ERROR' = 3 }
    if ($levelOrder[$Level] -lt $levelOrder[$script:LogConfig.LogLevel]) {
        return
    }

    $timestamp = Get-Date -Format 'o'
    $correlationId = $script:LogConfig.CorrelationId

    $logEntry = @{
        timestamp     = $timestamp
        level         = $Level
        message       = $Message
        correlationId = $correlationId
        properties    = $Properties
    }
    if ($Email) { $logEntry.email = $Email }

    # File logging
    if ($script:LogConfig.LogFilePath) {
        try {
            $logDir = Split-Path $script:LogConfig.LogFilePath -Parent
            if ($logDir -and -not (Test-Path $logDir)) {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }

            if ($script:LogConfig.JsonFormat) {
                $logLine = $logEntry | ConvertTo-Json -Compress
            }
            else {
                $logLine = "$timestamp [$Level] [$correlationId] $Message"
            }
            Add-Content -LiteralPath $script:LogConfig.LogFilePath -Value $logLine -Encoding UTF8
        }
        catch {
            Write-Warning "Log write failed: $($_.Exception.Message)"
        }
    }
}

# ============================================================================
# MODULE: Validation Functions
# ============================================================================

function Test-EmailAddress {
    <#
    .SYNOPSIS
        Validates an email address format
    .PARAMETER Email
        Email address to validate
    .OUTPUTS
        Returns validation result object with IsValid, Email, and Error properties
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    $result = @{
        IsValid = $false
        Email   = $Email.Trim()
        Error   = $null
    }

    $email = $Email.Trim()

    if ([string]::IsNullOrWhiteSpace($email)) {
        $result.Error = "Email is empty"
        return [PSCustomObject]$result
    }

    if ($email.Length -gt 254) {
        $result.Error = "Email exceeds 254 characters"
        return [PSCustomObject]$result
    }

    # Standard email regex
    $pattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

    if ($email -notmatch $pattern) {
        $result.Error = "Invalid email format"
        return [PSCustomObject]$result
    }

    $parts = $email -split '@'
    if ($parts[0].Length -gt 64) {
        $result.Error = "Local part exceeds 64 characters"
        return [PSCustomObject]$result
    }

    $result.IsValid = $true
    return [PSCustomObject]$result
}

function Test-UrlSafety {
    <#
    .SYNOPSIS
        Validates a URL for safety (blocks javascript:, data:, etc.)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Url
    )

    $result = @{
        IsValid = $false
        Url     = $Url
        Error   = $null
    }

    if ([string]::IsNullOrWhiteSpace($Url)) {
        $result.Error = "URL is empty"
        return [PSCustomObject]$result
    }

    $dangerousSchemes = @('javascript:', 'vbscript:', 'data:', 'file:', 'about:')
    $lowerUrl = $Url.ToLower().Trim()

    foreach ($scheme in $dangerousSchemes) {
        if ($lowerUrl.StartsWith($scheme)) {
            $result.Error = "Dangerous URL scheme: $scheme"
            return [PSCustomObject]$result
        }
    }

    if (-not ($lowerUrl.StartsWith('http://') -or $lowerUrl.StartsWith('https://'))) {
        $result.Error = "URL must use http:// or https://"
        return [PSCustomObject]$result
    }

    try {
        $uri = [System.Uri]::new($Url)
        if ([string]::IsNullOrWhiteSpace($uri.Host)) {
            $result.Error = "URL has no valid host"
            return [PSCustomObject]$result
        }
    }
    catch {
        $result.Error = "Malformed URL"
        return [PSCustomObject]$result
    }

    $result.IsValid = $true
    return [PSCustomObject]$result
}

# ============================================================================
# MODULE: Failed Recipients Tracking
# ============================================================================

function Add-FailedRecipient {
    <#
    .SYNOPSIS
        Records a failed email recipient to the failure log
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email,

        [Parameter(Mandatory = $true)]
        [string]$Reason,

        [Parameter(Mandatory = $true)]
        [string]$FailedRecipientsPath,

        [Parameter(Mandatory = $false)]
        [int]$AttemptCount = 0
    )

    try {
        $failDir = Split-Path $FailedRecipientsPath -Parent
        if ($failDir -and -not (Test-Path $failDir)) {
            New-Item -Path $failDir -ItemType Directory -Force | Out-Null
        }

        if (-not (Test-Path $FailedRecipientsPath)) {
            "timestamp;email;reason;attempts;correlationId" | Out-File -FilePath $FailedRecipientsPath -Encoding UTF8
        }

        $timestamp = Get-Date -Format 'o'
        $correlationId = $script:LogConfig.CorrelationId
        $escapedReason = $Reason -replace ';', ','

        $line = "$timestamp;$Email;$escapedReason;$AttemptCount;$correlationId"
        Add-Content -LiteralPath $FailedRecipientsPath -Value $line -Encoding UTF8

        Write-DigestLog -Level WARN -Message "Failed recipient recorded" -Email $Email -Properties @{
            reason = $Reason
            attempts = $AttemptCount
        }
    }
    catch {
        Write-Warning "Failed to write to failed recipients file: $($_.Exception.Message)"
    }
}

# ============================================================================
# MODULE: Email Authentication
# ============================================================================

function Get-EmailAuthenticationCredential {
    <#
    .SYNOPSIS
        Creates SMTP credential object based on authentication method
    .PARAMETER AuthMethod
        Authentication method: 'Basic' or 'OAuth'
    .PARAMETER Username
        SMTP username (email address)
    .PARAMETER Password
        Password for Basic auth
    .PARAMETER TenantId
        Azure AD Tenant ID for OAuth
    .PARAMETER ClientId
        Azure AD Application (Client) ID for OAuth
    .PARAMETER ClientSecret
        Azure AD Application Secret for OAuth
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Basic', 'OAuth')]
        [string]$AuthMethod,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(ParameterSetName = 'Basic', Mandatory = $true)]
        [string]$Password,

        [Parameter(ParameterSetName = 'OAuth', Mandatory = $true)]
        [string]$TenantId,

        [Parameter(ParameterSetName = 'OAuth', Mandatory = $true)]
        [string]$ClientId,

        [Parameter(ParameterSetName = 'OAuth', Mandatory = $true)]
        [string]$ClientSecret
    )

    switch ($AuthMethod) {
        'Basic' {
            Write-DigestLog -Level DEBUG -Message "Creating Basic Authentication credential" -Properties @{ username = $Username }
            Write-Verbose "Creating Basic Authentication credential for $Username"
            $securePass = ConvertTo-SecureString $Password -AsPlainText -Force
            return New-Object System.Management.Automation.PSCredential($Username, $securePass)
        }

        'OAuth' {
            Write-DigestLog -Level DEBUG -Message "Acquiring OAuth2 token" -Properties @{ username = $Username; tenantId = $TenantId }
            Write-Verbose "Acquiring OAuth2 token for $Username"
            try {
                $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

                $body = @{
                    client_id     = $ClientId
                    client_secret = $ClientSecret
                    scope         = "https://outlook.office365.com/.default"
                    grant_type    = "client_credentials"
                }

                $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"

                $oauthToken = $response.access_token
                $secureToken = ConvertTo-SecureString $oauthToken -AsPlainText -Force

                Write-DigestLog -Level INFO -Message "OAuth token acquired" -Properties @{ expiresIn = $response.expires_in }
                Write-Verbose "OAuth token acquired successfully"
                return New-Object System.Management.Automation.PSCredential($Username, $secureToken)
            }
            catch {
                $errorMsg = "Failed to acquire OAuth token: $($_.Exception.Message)"
                Write-DigestLog -Level ERROR -Message $errorMsg
                throw $errorMsg
            }
        }
    }
}

# ============================================================================
# MODULE: HTML Template Processing
# ============================================================================

function Get-ProcessedHtmlTemplate {
    <#
    .SYNOPSIS
        Loads and processes HTML template with placeholder replacement
    .PARAMETER TemplatePath
        Path to HTML template file
    .PARAMETER Replacements
        Hashtable of placeholder-value pairs for replacement
    .PARAMETER Encoding
        File encoding (default: UTF8)
    .PARAMETER ValidateUrls
        Validate URLs in replacements (default: $true)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ })]
        [string]$TemplatePath,

        [Parameter(Mandatory = $false)]
        [hashtable]$Replacements = @{},

        [Parameter(Mandatory = $false)]
        [string]$Encoding = 'UTF8',

        [Parameter(Mandatory = $false)]
        [bool]$ValidateUrls = $true
    )

    try {
        Write-DigestLog -Level DEBUG -Message "Loading HTML template" -Properties @{ path = $TemplatePath }
        Write-Verbose "Loading HTML template from: $TemplatePath"
        $htmlContent = Get-Content -LiteralPath $TemplatePath -Raw -Encoding $Encoding

        if (-not ([System.Management.Automation.PSTypeName]'System.Web.HttpUtility').Type) {
            Add-Type -AssemblyName System.Web
        }

        foreach ($key in $Replacements.Keys) {
            $value = $Replacements[$key]

            # URL placeholders - validate but don't HTML encode
            if ($key -match 'LINK|URL|HREF') {
                if ($ValidateUrls) {
                    $urlCheck = Test-UrlSafety -Url $value
                    if (-not $urlCheck.IsValid) {
                        Write-DigestLog -Level WARN -Message "Invalid URL replaced with #" -Properties @{
                            placeholder = $key
                            error = $urlCheck.Error
                        }
                        $value = "#"
                    }
                }
                $htmlContent = $htmlContent.Replace($key, $value)
            }
            else {
                # HTML encode non-URL values
                $encodedValue = [System.Web.HttpUtility]::HtmlEncode($value)
                $htmlContent = $htmlContent.Replace($key, $encodedValue)
            }
        }

        Write-DigestLog -Level DEBUG -Message "Template processed" -Properties @{ replacements = $Replacements.Count }
        Write-Verbose "Template processed with $($Replacements.Count) replacements"
        return $htmlContent
    }
    catch {
        $errorMsg = "Failed to process HTML template: $($_.Exception.Message)"
        Write-DigestLog -Level ERROR -Message $errorMsg
        throw $errorMsg
    }
}

# ============================================================================
# MODULE: Inline Image Handling
# ============================================================================

function New-EmailAlternateViewWithImages {
    <#
    .SYNOPSIS
        Creates AlternateView with embedded inline images
    .PARAMETER HtmlBody
        HTML content as string
    .PARAMETER InlineImages
        Array of hashtables with ContentId and FilePath
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter(Mandatory = $false)]
        [hashtable[]]$InlineImages = @()
    )

    try {
        $altView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString(
            $HtmlBody,
            [System.Text.Encoding]::UTF8,
            "text/html"
        )

        foreach ($image in $InlineImages) {
            if (-not (Test-Path $image.FilePath)) {
                Write-DigestLog -Level WARN -Message "Inline image not found" -Properties @{
                    contentId = $image.ContentId
                    path = $image.FilePath
                }
                Write-Warning "Inline image not found: $($image.FilePath)"
                continue
            }

            $linkedResource = New-Object System.Net.Mail.LinkedResource($image.FilePath)

            $extension = [System.IO.Path]::GetExtension($image.FilePath).TrimStart('.').ToLower()
            if ($extension -eq 'jpg') { $extension = 'jpeg' }

            $linkedResource.ContentType = New-Object System.Net.Mime.ContentType("image/$extension")
            $linkedResource.ContentId = $image.ContentId
            $linkedResource.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64

            [void]$altView.LinkedResources.Add($linkedResource)
            Write-DigestLog -Level DEBUG -Message "Added inline image" -Properties @{ contentId = $image.ContentId }
            Write-Verbose "Added inline image: $($image.ContentId)"
        }

        return $altView
    }
    catch {
        $errorMsg = "Failed to create AlternateView: $($_.Exception.Message)"
        Write-DigestLog -Level ERROR -Message $errorMsg
        throw $errorMsg
    }
}

# ============================================================================
# MODULE: SMTP Sending
# ============================================================================

function Send-HtmlEmail {
    <#
    .SYNOPSIS
        Sends HTML email with inline images and attachments
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$To,

        [Parameter(Mandatory = $true)]
        [string]$From,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter(Mandatory = $false)]
        [hashtable[]]$InlineImages = @(),

        [Parameter(Mandatory = $false)]
        [string[]]$Attachments = @(),

        [Parameter(Mandatory = $false)]
        [string]$Bcc,

        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [int]$SmtpPort,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory = $false)]
        [bool]$EnableSsl = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ValidateRecipient = $true
    )

    # Validate recipient
    if ($ValidateRecipient) {
        $validation = Test-EmailAddress -Email $To
        if (-not $validation.IsValid) {
            Write-DigestLog -Level WARN -Message "Invalid recipient" -Email $To -Properties @{ error = $validation.Error }
            return @{
                Success = $false
                Error   = "Invalid email: $($validation.Error)"
                Email   = $To
            }
        }
        $To = $validation.Email
    }

    $mailMessage = $null
    $smtpClient = $null

    try {
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $From
        $mailMessage.To.Add($To)
        if ($Bcc) { $mailMessage.Bcc.Add($Bcc) }
        $mailMessage.Subject = $Subject
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.IsBodyHtml = $true

        $altView = New-EmailAlternateViewWithImages -HtmlBody $HtmlBody -InlineImages $InlineImages
        [void]$mailMessage.AlternateViews.Add($altView)

        foreach ($attachmentPath in $Attachments) {
            if (Test-Path $attachmentPath) {
                $attachment = New-Object System.Net.Mail.Attachment($attachmentPath)
                [void]$mailMessage.Attachments.Add($attachment)
                Write-Verbose "Added attachment: $attachmentPath"
            }
            else {
                Write-Warning "Attachment not found: $attachmentPath"
            }
        }

        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtpClient.EnableSsl = $EnableSsl
        $smtpClient.Credentials = $Credential

        Write-DigestLog -Level DEBUG -Message "Sending email" -Email $To
        Write-Verbose "Sending email to: $To"
        $smtpClient.Send($mailMessage)
        Write-DigestLog -Level INFO -Message "Email sent" -Email $To
        Write-Verbose "Email sent successfully to: $To"

        return @{
            Success = $true
            Error   = $null
            Email   = $To
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-DigestLog -Level WARN -Message "Send failed" -Email $To -Properties @{ error = $errorMsg }
        Write-Warning "Failed to send to ${To}: $errorMsg"
        return @{
            Success = $false
            Error   = $errorMsg
            Email   = $To
        }
    }
    finally {
        if ($mailMessage) { $mailMessage.Dispose() }
        if ($smtpClient) { $smtpClient.Dispose() }
    }
}

# ============================================================================
# MODULE: Batch Sending with Checkpointing
# ============================================================================

function Send-BulkHtmlEmail {
    <#
    .SYNOPSIS
        Sends bulk HTML emails with batching, rate limiting, and checkpointing
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Recipients,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateConfig,

        [Parameter(Mandatory = $true)]
        [hashtable]$SmtpConfig,

        [Parameter(Mandatory = $false)]
        [int]$BatchSize = 20,

        [Parameter(Mandatory = $false)]
        [double]$WindowMinutes = 3.0,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [string]$CheckpointPath = "email_checkpoint.txt",

        [Parameter(Mandatory = $false)]
        [string]$FailedRecipientsPath = "failed_recipients.csv",

        [Parameter(Mandatory = $false)]
        [string]$LogFilePath
    )

    # Initialize logging
    $correlationId = [guid]::NewGuid().ToString('N').Substring(0, 8)
    Set-DigestLogConfig -LogFilePath $LogFilePath -CorrelationId $correlationId

    Write-DigestLog -Level INFO -Message "Starting bulk campaign" -Properties @{
        totalRecipients = $Recipients.Count
        batchSize = $BatchSize
    }

    # Load checkpoint
    $sentEmails = New-Object System.Collections.Generic.HashSet[string]
    if (Test-Path $CheckpointPath) {
        Get-Content $CheckpointPath | ForEach-Object {
            [void]$sentEmails.Add($_.Trim().ToLower())
        }
        Write-DigestLog -Level INFO -Message "Checkpoint loaded" -Properties @{ alreadySent = $sentEmails.Count }
        Write-Host "Loaded checkpoint: $($sentEmails.Count) already sent"
    }

    # Filter and validate recipients
    $pendingRecipients = @()
    $skippedInvalid = 0

    foreach ($recipient in $Recipients) {
        $email = $recipient.Email.Trim()

        if ([string]::IsNullOrWhiteSpace($email)) { continue }
        if ($sentEmails.Contains($email.ToLower())) { continue }

        $validation = Test-EmailAddress -Email $email
        if (-not $validation.IsValid) {
            $skippedInvalid++
            Add-FailedRecipient -Email $email -Reason "Invalid: $($validation.Error)" -FailedRecipientsPath $FailedRecipientsPath
            continue
        }

        $pendingRecipients += $recipient
    }

    if ($skippedInvalid -gt 0) {
        Write-DigestLog -Level WARN -Message "Invalid emails skipped" -Properties @{ count = $skippedInvalid }
    }

    if ($pendingRecipients.Count -eq 0) {
        Write-DigestLog -Level INFO -Message "No pending emails"
        Write-Host "No pending emails to send"
        return @{
            CampaignId = $correlationId
            Sent = 0
            Failed = $skippedInvalid
        }
    }

    Write-Host "`n=== BULK EMAIL CAMPAIGN ===" -ForegroundColor Cyan
    Write-Host "Campaign ID: $correlationId"
    Write-Host "Total pending: $($pendingRecipients.Count)"
    Write-Host "Already sent: $($sentEmails.Count)"
    Write-Host "Batch size: $BatchSize"
    Write-Host "Window: $WindowMinutes min"
    Write-Host "===========================`n" -ForegroundColor Cyan

    # Load template
    $baseHtml = Get-Content -LiteralPath $TemplateConfig.TemplatePath -Raw -Encoding $TemplateConfig.Encoding

    # Ensure System.Web is loaded
    if (-not ([System.Management.Automation.PSTypeName]'System.Web.HttpUtility').Type) {
        Add-Type -AssemblyName System.Web
    }

    $stats = @{ Sent = 0; Failed = 0; Retried = 0 }
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Process batches
    for ($offset = 0; $offset -lt $pendingRecipients.Count; $offset += $BatchSize) {
        $windowStart = $stopwatch.Elapsed
        $endIndex = [Math]::Min($offset + $BatchSize - 1, $pendingRecipients.Count - 1)
        $batch = $pendingRecipients[$offset..$endIndex]

        $batchNum = [Math]::Floor($offset / $BatchSize) + 1
        $totalBatches = [Math]::Ceiling($pendingRecipients.Count / $BatchSize)

        Write-Host ("[{0:HH:mm:ss}] === Batch {1}/{2} ===" -f (Get-Date), $batchNum, $totalBatches) -ForegroundColor Yellow

        foreach ($recipient in $batch) {
            $email = $recipient.Email.Trim()

            # Process template
            $htmlBody = $baseHtml
            foreach ($key in $recipient.Replacements.Keys) {
                $value = $recipient.Replacements[$key]

                if ($key -match 'LINK|URL|HREF') {
                    $urlCheck = Test-UrlSafety -Url $value
                    if (-not $urlCheck.IsValid) { $value = "#" }
                    $htmlBody = $htmlBody.Replace($key, $value)
                }
                else {
                    $encodedValue = [System.Web.HttpUtility]::HtmlEncode($value)
                    $htmlBody = $htmlBody.Replace($key, $encodedValue)
                }
            }

            # Retry loop
            $attempt = 0
            $success = $false
            $retryDelay = 2
            $lastError = $null

            while ($attempt -lt $MaxRetries -and -not $success) {
                $attempt++

                $sendParams = @{
                    To               = $email
                    From             = $SmtpConfig.From
                    Subject          = $SmtpConfig.Subject
                    HtmlBody         = $htmlBody
                    InlineImages     = $TemplateConfig.InlineImages
                    Attachments      = $TemplateConfig.Attachments
                    Bcc              = $SmtpConfig.Bcc
                    SmtpServer       = $SmtpConfig.Server
                    SmtpPort         = $SmtpConfig.Port
                    Credential       = $SmtpConfig.Credential
                    EnableSsl        = $SmtpConfig.EnableSsl
                    ValidateRecipient = $false
                }

                $result = Send-HtmlEmail @sendParams

                if ($result.Success) {
                    $success = $true
                    $stats.Sent++

                    try {
                        Add-Content -LiteralPath $CheckpointPath -Value $email -Encoding UTF8
                    }
                    catch {
                        Write-DigestLog -Level WARN -Message "Checkpoint write failed" -Email $email
                    }

                    Write-Host "  [OK] $email" -ForegroundColor Green
                }
                else {
                    $lastError = $result.Error
                    if ($attempt -lt $MaxRetries) {
                        $stats.Retried++
                        Write-Host "  [RETRY $attempt/$MaxRetries] $email" -ForegroundColor Yellow
                        Start-Sleep -Seconds $retryDelay
                        $retryDelay = [Math]::Min($retryDelay * 2, 30)
                    }
                }
            }

            if (-not $success) {
                $stats.Failed++
                Write-Host "  [FAIL] $email" -ForegroundColor Red
                Add-FailedRecipient -Email $email -Reason $lastError -FailedRecipientsPath $FailedRecipientsPath -AttemptCount $attempt
            }

            Start-Sleep -Milliseconds (Get-Random -Minimum 150 -Maximum 500)
        }

        # Window spacing
        if ($offset + $BatchSize -lt $pendingRecipients.Count) {
            $elapsed = $stopwatch.Elapsed - $windowStart
            $targetWindow = [TimeSpan]::FromMinutes($WindowMinutes)

            if ($elapsed -lt $targetWindow) {
                $sleepSeconds = [int](($targetWindow - $elapsed).TotalSeconds)
                if ($sleepSeconds -gt 0) {
                    Write-Host "`n[WAIT] ${sleepSeconds}s until next batch..." -ForegroundColor Cyan
                    Start-Sleep -Seconds $sleepSeconds
                }
            }
        }
        Write-Host ""
    }

    $stopwatch.Stop()

    Write-DigestLog -Level INFO -Message "Campaign complete" -Properties @{
        sent = $stats.Sent
        failed = $stats.Failed
        duration = $stopwatch.Elapsed.ToString('hh\:mm\:ss')
    }

    Write-Host "`n=== CAMPAIGN COMPLETE ===" -ForegroundColor Green
    Write-Host "Campaign ID: $correlationId"
    Write-Host "Duration: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
    Write-Host "Sent: $($stats.Sent)" -ForegroundColor Green
    Write-Host "Failed: $($stats.Failed)" -ForegroundColor $(if ($stats.Failed -gt 0) { 'Red' } else { 'Green' })
    if ($stats.Failed -gt 0) {
        Write-Host "Failures logged to: $FailedRecipientsPath" -ForegroundColor Yellow
    }
    Write-Host "========================`n" -ForegroundColor Green

    return @{
        CampaignId = $correlationId
        Sent = $stats.Sent
        Failed = $stats.Failed
        Retried = $stats.Retried
        Duration = $stopwatch.Elapsed
    }
}

# ============================================================================
# EXPORTS
# ============================================================================

Export-ModuleMember -Function @(
    'Set-DigestLogConfig',
    'Write-DigestLog',
    'Test-EmailAddress',
    'Test-UrlSafety',
    'Get-EmailAuthenticationCredential',
    'Get-ProcessedHtmlTemplate',
    'New-EmailAlternateViewWithImages',
    'Send-HtmlEmail',
    'Send-BulkHtmlEmail',
    'Add-FailedRecipient'
)
