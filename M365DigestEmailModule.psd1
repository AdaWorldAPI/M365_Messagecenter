@{
    # Module manifest for M365DigestEmailModule

    # Script module file associated with this manifest
    RootModule = 'M365DigestEmailModule.psm1'

    # Version number
    ModuleVersion = '1.1.0'

    # Unique ID for this module
    GUID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'

    # Author
    Author = 'Jan Huebener'

    # Company or vendor
    CompanyName = 'DATAGROUP'

    # Copyright statement
    Copyright = '(c) 2025. All rights reserved.'

    # Description
    Description = 'M365 Digest Email Sender - Professional HTML email system with OAuth/Basic Auth, inline images, batching, and checkpoint-based resume capability.'

    # Minimum PowerShell version required
    PowerShellVersion = '5.1'

    # Required .NET Framework version
    DotNetFrameworkVersion = '4.5'

    # Functions to export from this module
    FunctionsToExport = @(
        # Logging
        'Set-DigestLogConfig',
        'Write-DigestLog',

        # Validation
        'Test-EmailAddress',
        'Test-UrlSafety',

        # Authentication
        'Get-EmailAuthenticationCredential',

        # Template Processing
        'Get-ProcessedHtmlTemplate',

        # Email Construction
        'New-EmailAlternateViewWithImages',

        # Sending
        'Send-HtmlEmail',
        'Send-BulkHtmlEmail',

        # Tracking
        'Add-FailedRecipient'
    )

    # Cmdlets to export (none - this is a script module)
    CmdletsToExport = @()

    # Variables to export (none)
    VariablesToExport = @()

    # Aliases to export (none)
    AliasesToExport = @()

    # Private data
    PrivateData = @{
        PSData = @{
            # Tags for PSGallery
            Tags = @('Email', 'SMTP', 'M365', 'OAuth', 'HTML', 'Bulk', 'Newsletter')

            # License URI
            LicenseUri = ''

            # Project URI
            ProjectUri = 'https://github.com/PSScript/M365_Messagecenter'

            # Release notes
            ReleaseNotes = @'
## Version 1.1.0 (2025-12-15)

### New Features
- Added structured JSON logging with correlation IDs
- Added email validation (RFC 5322 compliant)
- Added URL safety validation (blocks javascript:, data:, etc.)
- Added failed recipients tracking (CSV log file)
- Added OAuth2 authentication script (Send-M365Digest-OAuth.ps1)

### Security Improvements
- Removed hardcoded credentials from scripts
- Added multiple secure credential loading options:
  - Interactive prompt (default)
  - Environment variables
  - Windows Credential Manager
  - DPAPI-encrypted file storage
- URLs in templates are now validated for safety
- Non-URL values are HTML-encoded to prevent XSS

### Breaking Changes
- Send-HtmlEmail now returns a hashtable instead of boolean
- BasicAuth script now requires credential method parameter

## Version 1.0.0 (2025-11-07)
- Initial release
- Basic and OAuth authentication
- Inline images with AlternateView
- Batch processing with checkpoints
- Modular architecture
'@
        }
    }
}
