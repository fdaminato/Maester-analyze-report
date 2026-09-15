<#
.SYNOPSIS
    Runs ITI365 and creates a brand-new standalone HTML security dashboard.

.DESCRIPTION
    1. Imports the ITI365 module.
    2. Restarts in a clean PowerShell process.
    3. Force-disconnects cached Graph, Exchange, Compliance, Teams and Azure contexts
       in isolated child processes so authentication assemblies cannot conflict.
    4. Clears previous contexts, opens a new interactive browser sign-in, and validates
       both the selected account and tenant before any ITI365 tests can run.
    5. Connects Graph first, then Exchange, Compliance, Teams and Azure using
       conflict-safe ordering and process-scoped contexts.
    5. Runs Invoke-ITI365 against the specified test folder with -SkipGraphConnect.
    6. Preserves ITI365's original HTML, JSON and Markdown reports.
    7. Removes the Markdown sections 'Remediation action' and 'Related links'
       from descriptions and test results.
    8. Builds a modern, searchable, sortable HTML dashboard with multi-select
       checkbox filters from the structured results returned by the same ITI365 run.
    9. Can rebuild the dashboard from an existing ITI365-Raw.json without rerunning tests.

.NOTES
    Recommended: PowerShell 7.x.
    The generated dashboard is self-contained and does not require internet access.

.EXAMPLE
    .\Invoke-MaesterModernDashboard-InteractiveLogin.ps1 -TestsPath "C:\Maester\maester-tests"

.EXAMPLE
    .\Invoke-MaesterModernDashboard-InteractiveLogin.ps1 `
        -TestsPath "C:\Maester\maester-tests" `
        -OutputRoot "C:\Maester\Reports" `
        -UserPrincipalName "admin@contoso.com" `
        -TenantId "00000000-0000-0000-0000-000000000000" `
        -IncludeLongRunning `
        -IncludePreview `
        -OpenReport

.EXAMPLE
    # Rebuild only the modern dashboard from a completed Maester run.
    .\Invoke-MaesterModernDashboard-InteractiveLogin.ps1 `
        -ExistingRunFolder "C:\Maester-Reports\2026-07-15_091500" `
        -OpenReport
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$TestsPath = (Get-Location).Path,

    [Parameter()]
    [string]$OutputRoot = (Join-Path -Path (Get-Location).Path -ChildPath 'maester-modern-reports'),

    [Parameter()]
    [string]$ExistingRunFolder,

    [Parameter()]
    [ValidateSet('None', 'Normal', 'Detailed', 'Diagnostic')]
    [string]$Verbosity = 'Normal',

    [Parameter()]
    [switch]$IncludeLongRunning,

    [Parameter()]
    [switch]$IncludePreview,

    [Parameter()]
    [switch]$SkipConnect,

    [Parameter()]
    [string]$UserPrincipalName,

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [switch]$SkipForcedDisconnect,

    [Parameter()]
    [switch]$SkipAzureConnection,

    [Parameter()]
    [switch]$RequireAllServices,

    [Parameter()]
    [switch]$OpenReport,

    [Parameter(DontShow)]
    [switch]$InternalCleanProcess
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:PowerShellExecutable = (Get-Process -Id $PID -ErrorAction Stop).Path

# Authentication modules such as Az.Accounts, Microsoft.Graph.Authentication,
# ExchangeOnlineManagement and MicrosoftTeams can load incompatible versions of
# Azure.Identity / MSAL into the same PowerShell process. A .NET assembly cannot
# be unloaded from a running process, so always restart this script in a clean,
# profile-free PowerShell process before importing any cloud modules.
if (-not $InternalCleanProcess) {
    $pwshPath = $script:PowerShellExecutable
    $childArguments = @(
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $PSCommandPath,
        '-InternalCleanProcess'
    )

    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Key -eq 'InternalCleanProcess') {
            continue
        }

        $parameterName = "-$($entry.Key)"
        if ($entry.Value -is [System.Management.Automation.SwitchParameter]) {
            if ($entry.Value.IsPresent) {
                $childArguments += $parameterName
            }
            continue
        }

        if ($null -ne $entry.Value) {
            $childArguments += $parameterName
            $childArguments += [string]$entry.Value
        }
    }

    Write-Host 'Restarting in a clean PowerShell process to prevent Azure.Identity/MSAL assembly conflicts...' -ForegroundColor Yellow
    & $pwshPath @childArguments
    exit $LASTEXITCODE
}

function Write-Step {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Get-LatestAvailableModule {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    return Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1
}

function Import-LatestModule {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [switch]$Required
    )

    $module = Get-LatestAvailableModule -Name $Name
    if ($null -eq $module) {
        if ($Required) {
            throw "Required PowerShell module '$Name' is not installed."
        }

        Write-Warning "Optional PowerShell module '$Name' is not installed. Related Maester tests will be skipped."
        return $null
    }

    Import-Module -Name $module.Path -Force -ErrorAction Stop
    Write-Host ("Loaded {0} {1}" -f $Name, $module.Version) -ForegroundColor DarkGreen
    return $module
}

function Invoke-ServiceConnection {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter()]
        [switch]$Required
    )

    try {
        & $ScriptBlock
        Write-Host "Connected: $Name" -ForegroundColor Green
        return $true
    }
    catch {
        $message = $_.Exception.Message
        if ($Required) {
            throw "Failed to connect to $Name. $message"
        }

        Write-Warning "Failed to connect to $Name. Related tests may be skipped. $message"
        return $false
    }
}

function Invoke-IsolatedSessionCleanup {
    <#
    .SYNOPSIS
        Runs one service cleanup in a separate profile-free PowerShell process.

    .DESCRIPTION
        Each cloud module can load its own Azure.Identity/MSAL assemblies. Running
        cleanup operations in separate processes prevents those assemblies from
        contaminating the main process that will connect Microsoft Graph first.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$ScriptContent
    )

    $temporaryScript = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath (
        'Maester-Disconnect-{0}-{1}.ps1' -f ($Name -replace '[^A-Za-z0-9_-]', '_'), ([guid]::NewGuid().ToString('N'))
    )

    try {
        [System.IO.File]::WriteAllText(
            $temporaryScript,
            $ScriptContent,
            [System.Text.UTF8Encoding]::new($false)
        )

        $cleanupOutput = & $script:PowerShellExecutable `
            -NoLogo `
            -NoProfile `
            -NonInteractive `
            -ExecutionPolicy Bypass `
            -File $temporaryScript 2>&1

        $cleanupExitCode = $LASTEXITCODE

        foreach ($line in @($cleanupOutput)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
                Write-Host ("  {0,-22} {1}" -f $Name, ([string]$line).Trim()) -ForegroundColor DarkGray
            }
        }

        if ($cleanupExitCode -ne 0) {
            Write-Warning "The isolated cleanup for $Name returned exit code $cleanupExitCode. Connection will still be attempted with a new process-scoped context."
            return $false
        }

        return $true
    }
    catch {
        Write-Warning "Unable to complete the isolated cleanup for $Name. $($_.Exception.Message)"
        return $false
    }
    finally {
        Remove-Item -Path $temporaryScript -Force -ErrorAction SilentlyContinue
    }
}

function Disconnect-AllMaesterServices {
    <#
    .SYNOPSIS
        Clears previous Maester-related cloud sessions before reconnecting.

    .DESCRIPTION
        Microsoft Graph and Azure can persist contexts across PowerShell processes.
        Exchange Online, Security & Compliance and Teams can also retain module
        connection state or cached tokens. Every cleanup runs in an isolated
        pwsh -NoProfile process so the dashboard's main process remains free of
        conflicting Azure.Identity/MSAL assemblies.
    #>

    Write-Step 'Forcing disconnect from all previous cloud sessions'

    $cleanupResults = [ordered]@{}

    $cleanupResults.Graph = Invoke-IsolatedSessionCleanup -Name 'Microsoft Graph' -ScriptContent @'
$ErrorActionPreference = 'Stop'
try {
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        Write-Output 'Module not installed; nothing to clear.'
        exit 0
    }

    Import-Module -Name $module.Path -Force -ErrorAction Stop
    $context = Get-MgContext -ErrorAction SilentlyContinue

    if ($null -ne $context) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Output 'Disconnected and cleared cached context.'
    }
    else {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Output 'No active context found; cached sign-in cleanup attempted.'
    }
}
catch {
    Write-Output ("Cleanup warning: {0}" -f $_.Exception.Message)
    exit 0
}
'@

    $cleanupResults.ExchangeCompliance = Invoke-IsolatedSessionCleanup -Name 'Exchange / Purview' -ScriptContent @'
$ErrorActionPreference = 'Stop'
try {
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        Write-Output 'Module not installed; nothing to clear.'
        exit 0
    }

    Import-Module -Name $module.Path -Force -ErrorAction Stop
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

    Get-PSSession -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ConfigurationName -match 'Microsoft.Exchange|Microsoft.Exchange.Management.ExoPowershellSnapin' -or
            $_.ComputerName -match 'outlook.office365.com|protection.outlook.com'
        } |
        Remove-PSSession -ErrorAction SilentlyContinue

    Write-Output 'Disconnected Exchange Online and Security & Compliance sessions.'
}
catch {
    Write-Output ("Cleanup warning: {0}" -f $_.Exception.Message)
    exit 0
}
'@

    $cleanupResults.Teams = Invoke-IsolatedSessionCleanup -Name 'Microsoft Teams' -ScriptContent @'
$ErrorActionPreference = 'Stop'
try {
    $module = Get-Module -ListAvailable -Name MicrosoftTeams |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        Write-Output 'Module not installed; nothing to clear.'
        exit 0
    }

    Import-Module -Name $module.Path -Force -ErrorAction Stop
    Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Output 'Disconnected Teams session.'
}
catch {
    Write-Output ("Cleanup warning: {0}" -f $_.Exception.Message)
    exit 0
}
'@

    $cleanupResults.Azure = Invoke-IsolatedSessionCleanup -Name 'Microsoft Azure' -ScriptContent @'
$ErrorActionPreference = 'Stop'
try {
    $module = Get-Module -ListAvailable -Name Az.Accounts |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        Write-Output 'Module not installed; nothing to clear.'
        exit 0
    }

    Import-Module -Name $module.Path -Force -ErrorAction Stop

    Disconnect-AzAccount -Scope Process -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Disconnect-AzAccount -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Clear-AzContext -Scope Process -Force -ErrorAction SilentlyContinue | Out-Null
    Clear-AzContext -Scope CurrentUser -Force -ErrorAction SilentlyContinue | Out-Null

    Write-Output 'Disconnected and removed saved Azure contexts.'
}
catch {
    Write-Output ("Cleanup warning: {0}" -f $_.Exception.Message)
    exit 0
}
'@

    Get-PSSession -ErrorAction SilentlyContinue |
        Remove-PSSession -ErrorAction SilentlyContinue

    Write-Host ''
    Write-Host 'Forced disconnect summary' -ForegroundColor Cyan
    foreach ($item in $cleanupResults.GetEnumerator()) {
        $status = if ($item.Value) { 'Completed' } else { 'Warning' }
        $color = if ($item.Value) { 'Green' } else { 'Yellow' }
        Write-Host ("  {0,-22} {1}" -f $item.Key, $status) -ForegroundColor $color
    }

    Write-Host 'Previous PowerShell SDK contexts were cleared. New interactive browser authentication will now be required.' -ForegroundColor Green
}

function Connect-MaesterServicesSafely {
    param(
        [Parameter()]
        [string]$UPN,

        [Parameter()]
        [string]$Tenant,

        [Parameter()]
        [switch]$SkipAzure,

        [Parameter()]
        [switch]$RequireAll
    )

    $connectionState = [ordered]@{
        Graph              = $false
        ExchangeOnline     = $false
        SecurityCompliance = $false
        Teams              = $false
        Azure              = $false
    }

    Import-LatestModule -Name 'Microsoft.Graph.Authentication' -Required | Out-Null
    Import-LatestModule -Name 'Maester' -Required | Out-Null

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host 'Microsoft Graph will use standard interactive authentication.' -ForegroundColor DarkGray

    Write-Step 'Connecting to Microsoft Graph first with a brand-new interactive sign-in'
    $connectionState.Graph = Invoke-ServiceConnection -Name 'Microsoft Graph' -Required -ScriptBlock {
        if ([string]::IsNullOrWhiteSpace($Tenant)) {
            throw 'A target TenantId is required for a fresh Microsoft Graph connection.'
        }
        if ([string]::IsNullOrWhiteSpace($UPN)) {
            throw 'A UserPrincipalName is required so the selected Graph account can be validated.'
        }

        $scopes = @(Get-MtGraphScope)
        if ($scopes.Count -eq 0) {
            throw 'Get-MtGraphScope returned no delegated permission scopes.'
        }

        $graphCommand = Get-Command Connect-MgGraph -ErrorAction Stop
        $lastValidationError = $null

        for ($attempt = 1; $attempt -le 3; $attempt++) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

            Write-Host ''
            Write-Host ("Graph authentication attempt {0}/3" -f $attempt) -ForegroundColor Yellow
            Write-Host ("  Required account : {0}" -f $UPN) -ForegroundColor Yellow
            Write-Host ("  Required tenant  : {0}" -f $Tenant) -ForegroundColor Yellow
            Write-Host '  An interactive sign-in window will open. Choose Use another account when necessary.' -ForegroundColor Yellow

            $graphParameters = @{
                Scopes       = $scopes
                TenantId     = $Tenant
                ContextScope = 'Process'
                NoWelcome    = $true
                ErrorAction  = 'Stop'
            }

            try {
                Connect-MgGraph @graphParameters
                $graphContext = Get-MgContext -ErrorAction Stop

                if ($null -eq $graphContext -or [string]::IsNullOrWhiteSpace([string]$graphContext.TenantId)) {
                    throw 'Connect-MgGraph completed, but no Microsoft Graph context was created.'
                }

                if (-not [string]::Equals(
                    [string]$graphContext.TenantId,
                    [string]$Tenant,
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    throw "Microsoft Graph connected to tenant '$($graphContext.TenantId)' instead of requested tenant '$Tenant'."
                }

                if (-not [string]::Equals(
                    [string]$graphContext.Account,
                    [string]$UPN,
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    throw "Microsoft Graph authenticated as '$($graphContext.Account)' instead of required account '$UPN'."
                }

                $organizationResponse = Invoke-MgGraphRequest `
                    -Method GET `
                    -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName' `
                    -ErrorAction Stop

                $organization = @($organizationResponse.value) | Select-Object -First 1
                if ($null -eq $organization -or [string]::IsNullOrWhiteSpace([string]$organization.id)) {
                    throw 'Microsoft Graph did not return an organization for the new session.'
                }

                if (-not [string]::Equals(
                    [string]$organization.id,
                    [string]$Tenant,
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {
                    throw "Graph organization '$($organization.id)' does not match requested tenant '$Tenant'."
                }

                Write-Host "Graph account      : $($graphContext.Account)" -ForegroundColor DarkGray
                Write-Host "Graph tenant       : $($graphContext.TenantId)" -ForegroundColor DarkGray
                Write-Host "Graph organization : $($organization.displayName)" -ForegroundColor DarkGray
                $lastValidationError = $null
                break
            }
            catch {
                $lastValidationError = $_.Exception.Message
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

                if ($attempt -lt 3) {
                    Write-Warning "$lastValidationError A new interactive sign-in window will be opened. Select the required account explicitly."
                    Start-Sleep -Seconds 2
                }
            }
        }

        if ($null -ne $lastValidationError) {
            throw "Unable to establish the requested fresh Microsoft Graph session after three attempts. $lastValidationError"
        }
    }

    Write-Step 'Connecting to Exchange Online'
    $exoModule = Import-LatestModule -Name 'ExchangeOnlineManagement'
    if ($null -ne $exoModule) {
        $connectionState.ExchangeOnline = Invoke-ServiceConnection -Name 'Exchange Online' -Required:$RequireAll -ScriptBlock {
            $exoParameters = @{
                ShowBanner  = $false
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($UPN)) {
                $exoParameters.UserPrincipalName = $UPN
            }

            $exoCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop
            if ($exoCommand.Parameters.ContainsKey('DisableWAM')) {
                $exoParameters.DisableWAM = $true
            }

            Connect-ExchangeOnline @exoParameters
        }

        if ($connectionState.ExchangeOnline) {
            if ([string]::IsNullOrWhiteSpace($UPN)) {
                $UPN = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                    Where-Object { $_.State -eq 'Connected' -and $_.UserPrincipalName } |
                    Select-Object -ExpandProperty UserPrincipalName -First 1
            }

            Write-Step 'Connecting to Security & Compliance PowerShell'
            $connectionState.SecurityCompliance = Invoke-ServiceConnection -Name 'Security & Compliance PowerShell' -Required:$RequireAll -ScriptBlock {
                $ippsParameters = @{
                    BypassMailboxAnchoring = $true
                    ShowBanner             = $false
                    ErrorAction            = 'Stop'
                }

                if (-not [string]::IsNullOrWhiteSpace($UPN)) {
                    $ippsParameters.UserPrincipalName = $UPN
                }

                $ippsCommand = Get-Command Connect-IPPSSession -ErrorAction Stop
                if ($ippsCommand.Parameters.ContainsKey('DisableWAM')) {
                    $ippsParameters.DisableWAM = $true
                }

                Connect-IPPSSession @ippsParameters
            }
        }
    }

    Write-Step 'Connecting to Microsoft Teams'
    $teamsModule = Import-LatestModule -Name 'MicrosoftTeams'
    if ($null -ne $teamsModule) {
        $connectionState.Teams = Invoke-ServiceConnection -Name 'Microsoft Teams' -Required:$RequireAll -ScriptBlock {
            $teamsParameters = @{
                ErrorAction = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
                $teamsParameters.TenantId = $Tenant
            }

            $teamsCommand = Get-Command Connect-MicrosoftTeams -ErrorAction Stop

            if (-not $teamsCommand.Parameters.ContainsKey('DisableWAM')) {
                throw "MicrosoftTeams $($teamsModule.Version) does not expose Connect-MicrosoftTeams -DisableWAM. Install MicrosoftTeams 7.8.1 or later."
            }

            Write-Host "Teams authentication: interactive browser, WAM disabled, no AccountId hint" -ForegroundColor DarkGray
            $teamsParameters.DisableWAM = $true

            $teamsConnection = Connect-MicrosoftTeams @teamsParameters

            if ($null -ne $teamsConnection -and -not [string]::IsNullOrWhiteSpace($Tenant)) {
                $teamsTenantProperty = $teamsConnection.PSObject.Properties['TenantId']
                if ($null -ne $teamsTenantProperty) {
                    $teamsTenantId = [string]$teamsTenantProperty.Value
                    if (-not [string]::IsNullOrWhiteSpace($teamsTenantId) -and
                        -not [string]::Equals($teamsTenantId, $Tenant, [System.StringComparison]::OrdinalIgnoreCase)) {
                        throw "Microsoft Teams connected to tenant '$teamsTenantId' instead of requested tenant '$Tenant'."
                    }
                }
            }
        }
    }

    if (-not $SkipAzure) {
        Write-Step 'Connecting to Azure last'
        $azModule = Import-LatestModule -Name 'Az.Accounts'
        if ($null -ne $azModule) {
            $connectionState.Azure = Invoke-ServiceConnection -Name 'Microsoft Azure' -Required:$RequireAll -ScriptBlock {
                $azParameters = @{
                    Scope                 = 'Process'
                    SkipContextPopulation = $true
                    ErrorAction           = 'Stop'
                }

                if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
                    $azParameters.Tenant = $Tenant
                }

                $azCommand = Get-Command Connect-AzAccount -ErrorAction Stop

                if (-not [string]::IsNullOrWhiteSpace($UPN) -and $azCommand.Parameters.ContainsKey('AccountId')) {
                    $azParameters.AccountId = $UPN
                }

                if ($azCommand.Parameters.ContainsKey('Force')) {
                    $azParameters.Force = $true
                }

                Connect-AzAccount @azParameters | Out-Null

                $azureContext = Get-AzContext -ErrorAction Stop
                if ($null -eq $azureContext) {
                    throw 'Connect-AzAccount completed, but no Azure context was created.'
                }

                if (-not [string]::IsNullOrWhiteSpace($Tenant) -and [string]$azureContext.Tenant.Id -ne $Tenant) {
                    throw "Microsoft Azure connected to tenant '$($azureContext.Tenant.Id)' instead of requested tenant '$Tenant'."
                }
            }
        }
    }
    else {
        Write-Warning 'Azure connection was skipped because -SkipAzureConnection was specified.'
    }

    Write-Host ''
    Write-Host 'Connection summary' -ForegroundColor Cyan
    foreach ($item in $connectionState.GetEnumerator()) {
        $color = if ($item.Value) { 'Green' } else { 'Yellow' }
        $status = if ($item.Value) { 'Connected' } else { 'Unavailable' }
        Write-Host ("  {0,-22} {1}" -f $item.Key, $status) -ForegroundColor $color
    }

    return [pscustomobject]$connectionState
}

function Get-ObjectValue {
    param(
        [Parameter()]
        $InputObject,

        [Parameter(Mandatory)]
        [string[]]$PropertyName,

        [Parameter()]
        $DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    foreach ($name in $PropertyName) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $DefaultValue
}

function ConvertTo-PlainText {
    param(
        [Parameter()]
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return $Value.Trim()
    }

    if ($Value -is [System.Management.Automation.ErrorRecord]) {
        return ($Value | Out-String).Trim()
    }

    if ($Value -is [System.Collections.IDictionary]) {
        try {
            return ($Value | ConvertTo-Json -Depth 8 -Compress -WarningAction SilentlyContinue)
        }
        catch {
            return ($Value | Out-String).Trim()
        }
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        $items = foreach ($item in $Value) {
            if ($null -ne $item) {
                [string]$item
            }
        }
        return ($items -join '; ').Trim()
    }

    return ([string]$Value).Trim()
}

function Remove-MaesterExcludedMarkdownSections {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $normalized = $Text `
        -replace '\\r\\n', "`n" `
        -replace '\\n', "`n" `
        -replace '\\r', "`n"

    $output = [System.Collections.Generic.List[string]]::new()
    $skipSection = $false
    $skipHeadingLevel = 0

    foreach ($rawLine in ($normalized -split "`n", 0, 'SimpleMatch')) {
        $line = $rawLine.TrimEnd("`r")
        $heading = [regex]::Match($line, '^\s*(#{1,6})\s+(.+?)\s*$')

        if ($skipSection) {
            if ($heading.Success -and $heading.Groups[1].Value.Length -le $skipHeadingLevel) {
                $skipSection = $false
                $skipHeadingLevel = 0
            }
            else {
                continue
            }
        }

        if ($heading.Success) {
            $headingText = $heading.Groups[2].Value.Trim()
            $headingText = $headingText.Trim([char[]]'*_` ')

            if ($headingText -match '^(?i:Remediation\s+action|Related\s+links)\s*:?\s*$') {
                $skipSection = $true
                $skipHeadingLevel = $heading.Groups[1].Value.Length
                continue
            }
        }

        $output.Add($line)
    }

    $cleaned = ($output -join "`n")
    $cleaned = [regex]::Replace($cleaned, '(?m)[ \t]+$', '')
    $cleaned = [regex]::Replace($cleaned, '(?:\r?\n){3,}', "`n`n")

    return $cleaned.Trim()
}

function Get-MaesterRequirementText {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$TestName
    )

    if ([string]::IsNullOrWhiteSpace($TestName)) {
        return 'The expected security configuration is present.'
    }

    $requirement = $TestName.Trim()

    $requirement = [regex]::Replace(
        $requirement,
        '^\s*[A-Z][A-Z0-9_-]{0,20}(?:\.[A-Z0-9_-]+)+\s*:\s*',
        '',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $requirement = [regex]::Replace($requirement, '^\s*Ensure\s+', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $requirement = [regex]::Replace($requirement, '^\s*Verify\s+', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $requirement = $requirement.Trim().TrimEnd('.')

    if ([string]::IsNullOrWhiteSpace($requirement)) {
        return 'The expected security configuration is present.'
    }

    if ($requirement.Length -gt 1 -and
        [char]::IsUpper($requirement[0]) -and
        -not [char]::IsUpper($requirement[1])) {
        $requirement = $requirement.Substring(0, 1).ToLowerInvariant() + $requirement.Substring(1)
    }

    return ($requirement + '.')
}

function Get-MaesterFallbackDescription {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$TestName
    )

    $requirement = Get-MaesterRequirementText -TestName $TestName
    return "This test verifies that $requirement"
}

function ConvertTo-FriendlyAssertionValue {
    param(
        [Parameter()]
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 'No value was returned'
    }

    $friendly = $Value.Trim()
    $friendly = $friendly.Trim([char[]]@([char]34, [char]39, [char]96))
    $friendly = $friendly -replace '^\$', ''

    switch -Regex ($friendly.ToLowerInvariant()) {
        '^true$'  { return 'True' }
        '^false$' { return 'False' }
        '^null$'  { return 'No value (null)' }
        '^empty$' { return 'Empty' }
        default   { return $friendly }
    }
}

function Get-MaesterErrorMessages {
    param(
        [Parameter()]
        $Value
    )

    $messages = [System.Collections.Generic.List[string]]::new()

    function Add-MessageValue {
        param($Candidate)

        if ($null -eq $Candidate) {
            return
        }

        try {
            $candidateText = ConvertTo-PlainText -Value $Candidate
        }
        catch {
            $candidateText = [string]$Candidate
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$candidateText)) {
            $candidateText = ([string]$candidateText).Trim()
            if (-not $messages.Contains($candidateText)) {
                [void]$messages.Add($candidateText)
            }
        }
    }

    foreach ($item in @($Value)) {
        if ($null -eq $item) {
            continue
        }

        try {
            if ($item -is [System.Management.Automation.ErrorRecord]) {
                if ($null -ne $item.Exception) {
                    Add-MessageValue $item.Exception.Message
                }

                if ($null -ne $item.ErrorDetails) {
                    Add-MessageValue $item.ErrorDetails.Message
                }

                if ($null -ne $item.TargetObject) {
                    $targetMessageProperty = $item.TargetObject.PSObject.Properties['Message']
                    if ($null -ne $targetMessageProperty) {
                        Add-MessageValue $targetMessageProperty.Value
                    }
                    elseif ($item.TargetObject -is [string]) {
                        Add-MessageValue $item.TargetObject
                    }
                }

                Add-MessageValue $item.ToString()
                continue
            }

            if ($item -is [string]) {
                Add-MessageValue $item
                continue
            }

            $exception = Get-ObjectValue -InputObject $item -PropertyName @('Exception')
            if ($null -ne $exception) {
                Add-MessageValue (Get-ObjectValue -InputObject $exception -PropertyName @('Message'))
            }

            $errorDetails = Get-ObjectValue -InputObject $item -PropertyName @('ErrorDetails')
            if ($null -ne $errorDetails) {
                Add-MessageValue (Get-ObjectValue -InputObject $errorDetails -PropertyName @('Message'))
            }

            $targetObject = Get-ObjectValue -InputObject $item -PropertyName @('TargetObject')
            if ($null -ne $targetObject) {
                $targetMessage = Get-ObjectValue -InputObject $targetObject -PropertyName @('Message')
                if ($null -ne $targetMessage) {
                    Add-MessageValue $targetMessage
                }
                elseif ($targetObject -is [string]) {
                    Add-MessageValue $targetObject
                }
            }

            Add-MessageValue (Get-ObjectValue -InputObject $item -PropertyName @('Message'))

            if ($messages.Count -eq 0) {
                Add-MessageValue $item
            }
        }
        catch {
            try {
                Add-MessageValue ([string]$item)
            }
            catch {
            }
        }
    }

    return @($messages)
}

function ConvertTo-ReadableMaesterError {
    param(
        [Parameter()]
        $Value,

        [Parameter()]
        [AllowNull()]
        [string]$TestName,

        [Parameter()]
        [AllowNull()]
        [string]$Status
    )

    if ($null -eq $Value) {
        return ''
    }

    $messages = @(Get-MaesterErrorMessages -Value $Value)
    if ($messages.Count -eq 0) {
        return ''
    }

    $raw = ($messages -join "`n`n")
    $raw = $raw `
        -replace '\\r\\n', "`n" `
        -replace '\\n', "`n" `
        -replace '\\r', "`n" `
        -replace '\\t', '  '

    $raw = Remove-MaesterExcludedMarkdownSections -Text $raw
    $requirement = Get-MaesterRequirementText -TestName $TestName
    $normalizedStatus = if ([string]::IsNullOrWhiteSpace($Status)) { 'failed' } else { $Status.ToLowerInvariant() }

    $assertion = [regex]::Match(
        $raw,
        '(?is)Expected\s+(?<expected>.+?)(?:,\s*because\s+(?<because>.+?))?,\s*but\s+got\s+(?<actual>.+?)(?:\.(?:\s|$)|\r?\n|$)'
    )

    if ($assertion.Success) {
        $expected = ConvertTo-FriendlyAssertionValue -Value $assertion.Groups['expected'].Value
        $actual = ConvertTo-FriendlyAssertionValue -Value $assertion.Groups['actual'].Value
        $because = $assertion.Groups['because'].Value.Trim().Trim([char[]]@([char]34, [char]39, [char]96)).TrimEnd('.')

        $expectedText = if (-not [string]::IsNullOrWhiteSpace($because)) {
            if ($because.Length -gt 1 -and [char]::IsLower($because[0])) {
                $because = $because.Substring(0, 1).ToUpperInvariant() + $because.Substring(1)
            }
            "$because."
        }
        else {
            $requirement
        }

        $outcomeText = switch ($normalizedStatus) {
            'error'   { 'Error: The test encountered an assertion error because the expected configuration was not detected.' }
            'skipped' { 'The test could not confirm the expected configuration.' }
            default   { 'Failed: The expected configuration was not detected.' }
        }

        return (@(
            '### Test outcome'
            $outcomeText
            ''
            '### Expected configuration'
            $expectedText
            ''
            '### Observed result'
            "The check returned $actual instead of $expected."
        ) -join "`n")
    }

    if ($raw -match '(?i)AadPremiumLicenseRequired|tenant needs to have Microsoft Entra ID P2|Microsoft Entra ID Governance license') {
        return (@(
            '### Test outcome'
            'This test could not be completed because the tenant does not have the required Microsoft Entra license.'
            ''
            '### Required license'
            'Microsoft Entra ID P2 or Microsoft Entra ID Governance.'
            ''
            '### Assessment note'
            'Treat this check as not assessed rather than as a confirmed configuration failure.'
        ) -join "`n")
    }

    $jsonMessage = [regex]::Match($raw, '(?is)"message"\s*:\s*"(?<message>(?:\\.|[^"\\])+)"')
    if ($jsonMessage.Success) {
        $apiMessage = $jsonMessage.Groups['message'].Value
        try {
            $apiMessage = ('"' + $apiMessage + '"' | ConvertFrom-Json -ErrorAction Stop)
        }
        catch {
            $apiMessage = $apiMessage -replace '\\"', '"' -replace '\\n', ' '
        }

        return (@(
            '### Test outcome'
            'The test could not complete its Microsoft 365 data query.'
            ''
            '### Service response'
            $apiMessage.Trim()
            ''
            '### Expected configuration'
            $requirement
        ) -join "`n")
    }

    $skipped = [regex]::Match($raw, '(?is)is\s+skipped,\s+because\s+(?<reason>.*?)(?:```|$)')
    if ($skipped.Success) {
        $reason = $skipped.Groups['reason'].Value.Trim()
        $reason = [regex]::Replace($reason, '(?im)^\s*(?:Invoke-[^:]+|Line\s*\||\d+\s*\||\|?\s*~+).*$','')
        $reason = [regex]::Replace($reason, '(?:\r?\n){3,}', "`n`n").Trim()

        if (-not [string]::IsNullOrWhiteSpace($reason)) {
            return (@(
                '### Test outcome'
                'The test was skipped.'
                ''
                '### Reason'
                $reason
                ''
                '### Expected configuration'
                $requirement
            ) -join "`n")
        }
    }

    $cleanLines = [System.Collections.Generic.List[string]]::new()
    foreach ($rawLine in ($raw -split "`n")) {
        $line = $rawLine.Trim()

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        if ($line -match '^(?i:InvalidResult|RuntimeException|Exception):\s*(?:[A-Z]:\\|/).+\.ps1\s*$') {
            continue
        }
        if ($line -match '^(?i:Line)\s*\|\s*$') {
            continue
        }
        if ($line -match '^\d+\s*\|') {
            continue
        }
        if ($line -match '^\|?\s*[~^]+\s*$') {
            continue
        }
        if ($line -match '^(?:[A-Z]:\\|/).+\.ps1(?::\d+)?$') {
            continue
        }
        if ($line -match '^```') {
            continue
        }

        $line = $line -replace '^\|\s*', ''
        if (-not [string]::IsNullOrWhiteSpace($line) -and -not $cleanLines.Contains($line)) {
            $cleanLines.Add($line)
        }
    }

    $cleanText = ($cleanLines -join "`n").Trim()
    if ([string]::IsNullOrWhiteSpace($cleanText)) {
        $cleanText = 'The test did not return a readable error message.'
    }

    $genericOutcome = switch ($normalizedStatus) {
        'error'   { 'Error: The test could not be completed and requires review.' }
        'skipped' { 'The test was skipped and requires review.' }
        'passed'  { 'The test completed successfully.' }
        default   { 'Failed: The test did not meet the expected condition.' }
    }

    return (@(
        '### Test outcome'
        $genericOutcome
        ''
        '### Result detail'
        $cleanText
        ''
        '### Expected configuration'
        $requirement
    ) -join "`n")
}

function ConvertTo-HtmlEncoded {
    param(
        [Parameter()]
        $Value
    )

    return [System.Net.WebUtility]::HtmlEncode((ConvertTo-PlainText -Value $Value))
}

function ConvertTo-NormalizedStatus {
    param(
        [Parameter()]
        [string]$Status,

        [Parameter()]
        [bool]$Investigate = $false
    )

    if ($Investigate) {
        return 'Investigate'
    }

    $value = if ($null -eq $Status) { '' } else { $Status.Trim().ToLowerInvariant() }

    switch -Regex ($value) {
        '^(pass|passed|success|succeeded)$'      { return 'Passed' }
        '^(fail|failed|failure)$'               { return 'Failed' }
        'investigat'                            { return 'Investigate' }
        '^(skip|skipped|ignored|inconclusive)$' { return 'Skipped' }
        '^(notrun|not run|not_run|pending)$'    { return 'NotRun' }
        '^(error|errored|broken)$'              { return 'Error' }
        default {
            if ([string]::IsNullOrWhiteSpace($Status)) {
                return 'NotRun'
            }
            return (Get-Culture).TextInfo.ToTitleCase($Status.Trim().ToLowerInvariant())
        }
    }
}

function ConvertTo-NormalizedSeverity {
    param(
        [Parameter()]
        [string]$Severity
    )

    $value = if ($null -eq $Severity) { '' } else { $Severity.Trim().ToLowerInvariant() }

    switch ($value) {
        'critical'      { return 'Critical' }
        'high'          { return 'High' }
        'medium'        { return 'Medium' }
        'low'           { return 'Low' }
        'info'          { return 'Info' }
        'informational' { return 'Info' }
        default         { return 'Not specified' }
    }
}

function ConvertTo-DurationText {
    param(
        [Parameter()]
        $Duration
    )

    if ($null -eq $Duration) {
        return ''
    }

    if ($Duration -is [TimeSpan]) {
        if ($Duration.TotalMinutes -ge 1) {
            return ('{0:0}m {1:00}s' -f [math]::Floor($Duration.TotalMinutes), $Duration.Seconds)
        }
        if ($Duration.TotalSeconds -ge 1) {
            return ('{0:0.00}s' -f $Duration.TotalSeconds)
        }
        return ('{0:0}ms' -f $Duration.TotalMilliseconds)
    }

    $milliseconds = 0.0
    if ([double]::TryParse([string]$Duration, [ref]$milliseconds)) {
        if ($milliseconds -ge 60000) {
            return ('{0:0.0}m' -f ($milliseconds / 60000))
        }
        if ($milliseconds -ge 1000) {
            return ('{0:0.00}s' -f ($milliseconds / 1000))
        }
        return ('{0:0}ms' -f $milliseconds)
    }

    return [string]$Duration
}

function Get-ControlId {
    param(
        [Parameter()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    $match = [regex]::Match(
        $Name,
        '\b[A-Z][A-Z0-9_-]{1,12}\.[A-Z0-9_-]+(?:\.[A-Z0-9_-]+)*\b',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($match.Success) {
        return $match.Value.ToUpperInvariant()
    }

    $colonMatch = [regex]::Match($Name, '^\s*([^:]{2,30}):')
    if ($colonMatch.Success) {
        return $colonMatch.Groups[1].Value.Trim()
    }

    return ''
}

function Get-CategoryName {
    param(
        [Parameter()]
        [string]$Service,

        [Parameter()]
        [string]$ControlId,

        [Parameter()]
        [string]$Tags
    )

    if (-not [string]::IsNullOrWhiteSpace($Service)) {
        return $Service.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($ControlId) -and $ControlId.Contains('.')) {
        return $ControlId.Split('.')[0].ToUpperInvariant()
    }

    if (-not [string]::IsNullOrWhiteSpace($Tags)) {
        return ($Tags -split '[,;]')[0].Trim()
    }

    return 'Other'
}

function Get-ResultCollection {
    param(
        [Parameter(Mandatory)]
        $MaesterResults
    )

    $tests = Get-ObjectValue -InputObject $MaesterResults -PropertyName @('Tests', 'TestResults', 'Results')

    if ($null -eq $tests -and ($MaesterResults -is [System.Collections.IEnumerable]) -and -not ($MaesterResults -is [string])) {
        $tests = $MaesterResults
    }

    return @($tests)
}

function New-NormalizedTestCollection {
    param(
        [Parameter(Mandatory)]
        $MaesterResults
    )

    $normalized = [System.Collections.Generic.List[object]]::new()
    $index = 0

    foreach ($test in (Get-ResultCollection -MaesterResults $MaesterResults)) {
        if ($null -eq $test) {
            continue
        }

        $index++

        try {
            $name = ConvertTo-PlainText (Get-ObjectValue -InputObject $test -PropertyName @(
                'Name', 'TestName', 'TestTitle', 'ExpandedName', 'Title'
            ) -DefaultValue "Test $index")

            $rawStatus = ConvertTo-PlainText (Get-ObjectValue -InputObject $test -PropertyName @(
                'Result', 'Status', 'Outcome'
            ) -DefaultValue 'NotRun')

            $investigateValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'Investigate', 'TestInvestigate', 'RequiresInvestigation'
            ) -DefaultValue $false

            $investigate = $false
            if ($investigateValue -is [bool]) {
                $investigate = $investigateValue
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$investigateValue)) {
                $investigate = ([string]$investigateValue -match '^(true|1|yes)$')
            }

            $status = ConvertTo-NormalizedStatus -Status $rawStatus -Investigate:$investigate

            $resultDetail = Get-ObjectValue -InputObject $test -PropertyName @(
                'ResultDetail', 'ResultDetails'
            )

            $severityValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'Severity', 'Risk', 'Level'
            )
            if ($null -eq $severityValue -or [string]::IsNullOrWhiteSpace((ConvertTo-PlainText $severityValue))) {
                $severityValue = Get-ObjectValue -InputObject $resultDetail -PropertyName @(
                    'Severity', 'Risk', 'Level'
                )
            }
            $severity = ConvertTo-NormalizedSeverity -Severity (ConvertTo-PlainText $severityValue)

            $serviceValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'Service', 'Workload', 'Product'
            )
            if ($null -eq $serviceValue -or [string]::IsNullOrWhiteSpace((ConvertTo-PlainText $serviceValue))) {
                $serviceValue = Get-ObjectValue -InputObject $resultDetail -PropertyName @(
                    'Service', 'Workload', 'Product'
                )
            }
            $service = ConvertTo-PlainText $serviceValue

            $tags = ConvertTo-PlainText (Get-ObjectValue -InputObject $test -PropertyName @(
                'Tag', 'Tags'
            ))

            $descriptionValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'TestDescription', 'Description', 'Overview'
            )
            if ($null -eq $descriptionValue -or [string]::IsNullOrWhiteSpace((ConvertTo-PlainText $descriptionValue))) {
                $descriptionValue = Get-ObjectValue -InputObject $resultDetail -PropertyName @(
                    'TestDescription', 'Description', 'Overview'
                )
            }
            $description = Remove-MaesterExcludedMarkdownSections -Text (ConvertTo-PlainText $descriptionValue)

            if ([string]::IsNullOrWhiteSpace($description)) {
                $description = Get-MaesterFallbackDescription -TestName $name
            }

            $detailsValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'TestResult', 'Details', 'Message', 'ResultMessage'
            )
            if ($null -eq $detailsValue -or [string]::IsNullOrWhiteSpace((ConvertTo-PlainText $detailsValue))) {
                $detailsValue = Get-ObjectValue -InputObject $resultDetail -PropertyName @(
                    'TestResult', 'Details', 'Message', 'ResultMessage'
                )
            }
            $details = Remove-MaesterExcludedMarkdownSections -Text (ConvertTo-PlainText $detailsValue)

            if (-not [string]::IsNullOrWhiteSpace($details) -and
                $details -match '(?is)InvalidResult:|Expected\s+.+?but\s+got|^\s*Line\s*\|') {
                $details = ConvertTo-ReadableMaesterError -Value $details -TestName $name -Status $status
            }

            $skipReasonValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'SkippedReason', 'SkipReason', 'Because'
            )
            if ($null -eq $skipReasonValue -or [string]::IsNullOrWhiteSpace((ConvertTo-PlainText $skipReasonValue))) {
                $skipReasonValue = Get-ObjectValue -InputObject $resultDetail -PropertyName @(
                    'SkippedReason', 'SkipReason', 'Because'
                )
            }
            $skipReason = Remove-MaesterExcludedMarkdownSections -Text (ConvertTo-PlainText $skipReasonValue)

            $errorValue = Get-ObjectValue -InputObject $test -PropertyName @(
                'ErrorRecord', 'Error', 'Exception'
            )
            $errorText = ConvertTo-ReadableMaesterError -Value $errorValue -TestName $name -Status $status

            if ([string]::IsNullOrWhiteSpace($details)) {
                if (-not [string]::IsNullOrWhiteSpace($skipReason)) {
                    $details = $skipReason
                }
                elseif (-not [string]::IsNullOrWhiteSpace($errorText)) {
                    $details = $errorText
                }
            }

            $duration = ConvertTo-DurationText (Get-ObjectValue -InputObject $test -PropertyName @(
                'Duration', 'ExecutionTime', 'Elapsed', 'Time'
            ))

            $path = ConvertTo-PlainText (Get-ObjectValue -InputObject $test -PropertyName @(
                'Path', 'File', 'Source', 'ScriptBlockFile'
            ))

            $helpUrl = ConvertTo-PlainText (Get-ObjectValue -InputObject $test -PropertyName @(
                'HelpUrl', 'DocumentationUrl', 'DocsUrl'
            ))

            $controlId = Get-ControlId -Name $name
            $category = Get-CategoryName -Service $service -ControlId $controlId -Tags $tags

            $normalized.Add([pscustomobject][ordered]@{
                Index       = $index
                ControlId   = $controlId
                Name        = $name
                Status      = $status
                Severity    = $severity
                Service     = $(if ([string]::IsNullOrWhiteSpace($service)) { $category } else { $service })
                Category    = $category
                Duration    = $duration
                Tags        = $tags
                Description = $description
                Details     = $details
                SkipReason  = $skipReason
                Error       = $errorText
                Source      = $path
                HelpUrl     = $helpUrl
            })
        }
        catch {
            $fallbackName = try {
                ConvertTo-PlainText (Get-ObjectValue -InputObject $test -PropertyName @('Name','TestName','TestTitle','ExpandedName','Title') -DefaultValue "Test $index")
            }
            catch {
                "Test $index"
            }

            $normalizationError = $_.Exception.Message
            $normalized.Add([pscustomobject][ordered]@{
                Index       = $index
                ControlId   = Get-ControlId -Name $fallbackName
                Name        = $fallbackName
                Status      = 'Error'
                Severity    = 'Unknown'
                Service     = 'Other'
                Category    = 'Other'
                Duration    = '00:00:00'
                Tags        = ''
                Description = Get-MaesterFallbackDescription -TestName $fallbackName
                Details     = "### Test outcome`nThis individual result could not be fully formatted.`n`n### Formatting detail`n$normalizationError"
                SkipReason  = ''
                Error       = $normalizationError
                Source      = ''
                HelpUrl     = ''
            })
        }
    }

    return $normalized
}

function Get-CountByProperty {
    param(
        [Parameter(Mandatory)]
        [object[]]$Items,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $result = [ordered]@{}

    foreach ($group in ($Items | Group-Object -Property $PropertyName | Sort-Object -Property Count -Descending)) {
        $name = if ([string]::IsNullOrWhiteSpace([string]$group.Name)) { 'Not specified' } else { [string]$group.Name }
        $result[$name] = $group.Count
    }

    return $result
}

function Get-RootDuration {
    param(
        [Parameter(Mandatory)]
        $MaesterResults
    )

    $duration = Get-ObjectValue -InputObject $MaesterResults -PropertyName @('Duration', 'TotalDuration', 'Elapsed')
    return ConvertTo-DurationText -Duration $duration
}

function ConvertTo-SafeJsonForHtml {
    param(
        [Parameter(Mandatory)]
        $InputObject
    )

    $json = $InputObject | ConvertTo-Json -Depth 12 -Compress -WarningAction SilentlyContinue
    return $json.Replace('</', '<\/').Replace('<!--', '<\!--')
}

function New-ModernDashboardHtml {
    param(
        [Parameter(Mandatory)]
        $MaesterResults,

        [Parameter(Mandatory)]
        [object[]]$Tests,

        [Parameter(Mandatory)]
        [string]$OriginalHtmlFileName,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $tenantName = ConvertTo-PlainText (Get-ObjectValue -InputObject $MaesterResults -PropertyName @('TenantName', 'OrganizationName') -DefaultValue 'Microsoft 365 tenant')
    $tenantId = ConvertTo-PlainText (Get-ObjectValue -InputObject $MaesterResults -PropertyName @('TenantId', 'OrganizationId') -DefaultValue '')
    $moduleVersion = ConvertTo-PlainText (Get-ObjectValue -InputObject $MaesterResults -PropertyName @('CurrentVersion', 'ModuleVersion', 'MaesterVersion') -DefaultValue '')
    $runDuration = Get-RootDuration -MaesterResults $MaesterResults
    $generatedAt = Get-Date

    $statusCounts = Get-CountByProperty -Items $Tests -PropertyName 'Status'
    $severityCounts = Get-CountByProperty -Items $Tests -PropertyName 'Severity'
    $serviceCounts = Get-CountByProperty -Items $Tests -PropertyName 'Service'

    $passed = @($Tests | Where-Object Status -eq 'Passed').Count
    $failed = @($Tests | Where-Object Status -eq 'Failed').Count
    $investigate = @($Tests | Where-Object Status -eq 'Investigate').Count
    $skipped = @($Tests | Where-Object Status -eq 'Skipped').Count
    $notRun = @($Tests | Where-Object Status -eq 'NotRun').Count
    $errors = @($Tests | Where-Object Status -eq 'Error').Count
    $total = $Tests.Count
    $actionRequired = $failed + $investigate + $errors
    $assessed = $passed + $failed + $investigate + $errors
    $score = if ($assessed -gt 0) { [math]::Round(($passed / $assessed) * 100, 1) } else { 0 }

    $dashboardData = [ordered]@{
        tenant = [ordered]@{
            name = $tenantName
            id = $tenantId
            moduleVersion = $moduleVersion
            generatedAt = $generatedAt.ToString('yyyy-MM-dd HH:mm:ss')
            runDuration = $runDuration
            originalReport = $OriginalHtmlFileName
        }
        summary = [ordered]@{
            total = $total
            passed = $passed
            failed = $failed
            investigate = $investigate
            skipped = $skipped
            notRun = $notRun
            error = $errors
            actionRequired = $actionRequired
            score = $score
        }
        statusCounts = $statusCounts
        severityCounts = $severityCounts
        serviceCounts = $serviceCounts
        tests = $Tests
    }

    $dataJson = ConvertTo-SafeJsonForHtml -InputObject $dashboardData

    $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>ITI365 - Security Assessment</title>
<style>
:root {
  --bg:#f4f7fb;
  --surface:#ffffff;
  --surface2:#eef3f9;
  --surface3:#dde7f2;
  --border:#e2eaf3;
  --text:#0f1e33;
  --muted:#5e7292;
  --accent:#2563eb;
  --navy:#1e3a8a;
  --green:#059669;
  --green-soft:#d1fae5;
  --red:#dc2626;
  --red-soft:#fde2e2;
  --amber:#d97706;
  --amber-soft:#fef3c7;
  --purple:#9333ea;
  --purple-soft:#f3e8ff;
  --gray:#64748b;
  --gray-soft:#e8edf4;
  --radius:12px;
  --radius-sm:8px;
  --shadow:0 2px 6px rgb(30 60 120 / .06), 0 1px 2px rgb(30 60 120 / .04);
  --shadow-hover:0 8px 22px rgb(30 60 120 / .13);
  --font:Inter, "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
}

[data-theme="dark"] {
  --bg:#0c1524;
  --surface:#14213a;
  --surface2:#1c2c49;
  --surface3:#284066;
  --border:#243a5e;
  --text:#e8f0fb;
  --muted:#93a7c4;
  --accent:#60a5fa;
  --navy:#93c5fd;
  --green:#34d399;
  --green-soft:rgba(16,185,129,.15);
  --red:#f87171;
  --red-soft:rgba(239,68,68,.15);
  --amber:#fbbf24;
  --amber-soft:rgba(245,158,11,.15);
  --purple:#c084fc;
  --purple-soft:rgba(168,85,247,.15);
  --gray:#a4b2c6;
  --gray-soft:rgba(148,163,184,.15);
  --shadow:0 2px 6px rgb(0 0 0 / .30);
  --shadow-hover:0 8px 22px rgb(0 0 0 / .45);
}

* { box-sizing:border-box; }
html { scroll-behavior:smooth; }

body {
  margin:0;
  min-height:100vh;
  background:var(--bg);
  color:var(--text);
  font:14px/1.45 var(--font);
}

button, input { font:inherit; }

header {
  position:sticky;
  top:0;
  z-index:100;
}

.topbar {
  min-height:68px;
  display:flex;
  align-items:center;
  gap:18px;
  padding:12px 28px;
  background:var(--surface);
  border-top:3px solid var(--accent);
  border-bottom:1px solid var(--border);
  box-shadow:var(--shadow);
}

.brand-left {
  min-width:0;
  display:flex;
  align-items:center;
  gap:12px;
}

.logo-fallback {
  width:42px;
  height:42px;
  min-width:42px;
  overflow:hidden;
  padding:0 7px;
  border-radius:10px;
  display:flex;
  align-items:center;
  justify-content:center;
  background:linear-gradient(135deg,var(--accent),var(--navy));
  color:#fff;
  font-size:16px;
  line-height:1.05;
  text-align:center;
  font-weight:800;
  box-shadow:var(--shadow);
}

h1 {
  margin:0;
  font-size:16px;
  line-height:1.2;
  letter-spacing:-.01em;
}

.subtitle {
  margin-top:3px;
  color:var(--muted);
  font-size:11.5px;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
  max-width:720px;
}

.topbar-actions {
  margin-left:auto;
  display:flex;
  align-items:center;
  gap:10px;
}

.btn {
  border:1px solid var(--border);
  border-radius:var(--radius-sm);
  padding:8px 12px;
  background:var(--surface);
  color:var(--text);
  cursor:pointer;
  font-size:12.5px;
  font-weight:600;
  white-space:nowrap;
  transition:.15s ease;
  text-decoration:none;
  display:inline-flex;
  align-items:center;
  gap:7px;
}

.btn:hover {
  background:var(--surface2);
  transform:translateY(-1px);
}

.btn-primary {
  background:var(--accent);
  color:#fff;
  border-color:transparent;
}

.btn-primary:hover {
  background:var(--navy);
}

.generated {
  min-width:164px;
  padding-left:12px;
  border-left:1px solid var(--border);
  color:var(--muted);
  font-size:11.5px;
  line-height:1.35;
  text-align:right;
}

.generated strong {
  color:var(--text);
  font-weight:700;
}

.layout {
  max-width:1800px;
  margin:0 auto;
  padding:26px 32px 38px;
}

.grid {
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
  gap:14px;
  margin-bottom:26px;
}

.grid > .card {
  position:relative;
  min-height:132px;
  padding:16px 17px;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  box-shadow:var(--shadow);
  overflow:hidden;
  transition:.15s ease;
}

.grid > .card:hover {
  box-shadow:var(--shadow-hover);
  transform:translateY(-1px);
}

.grid > .card::before {
  content:"";
  position:absolute;
  top:0;
  left:0;
  right:0;
  height:3px;
  background:var(--accent);
  opacity:.8;
}

.grid > .card.card-good::before { background:var(--green); }
.grid > .card.card-bad::before { background:var(--red); }
.grid > .card.card-warn::before { background:var(--amber); }
.grid > .card.card-investigate::before { background:var(--purple); }

.card-title {
  color:var(--muted);
  font-size:10.5px;
  font-weight:700;
  line-height:1.35;
  text-transform:uppercase;
  letter-spacing:.055em;
}

.card-value {
  margin-top:9px;
  color:var(--text);
  font-size:27px;
  font-weight:800;
  line-height:1;
  letter-spacing:-.025em;
}

.card-note {
  margin-top:7px;
  color:var(--muted);
  font-size:11.5px;
  line-height:1.4;
}

.good { color:var(--green) !important; }
.bad { color:var(--red) !important; }
.warn { color:var(--amber) !important; }
.info { color:var(--accent) !important; }
.investigate { color:var(--purple) !important; }

.section { margin-top:26px; }

.section h2 {
  display:flex;
  align-items:center;
  gap:8px;
  margin:0 0 13px;
  color:var(--muted);
  font-size:12.5px;
  font-weight:800;
  text-transform:uppercase;
  letter-spacing:.065em;
}

.section h2::before {
  content:"";
  width:4px;
  height:17px;
  border-radius:999px;
  background:var(--accent);
}

.mini-grid {
  display:grid;
  grid-template-columns:repeat(3,minmax(260px,1fr));
  gap:14px;
}

.section .card {
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  box-shadow:var(--shadow);
}

.chart-card {
  min-height:220px;
  display:grid;
  grid-template-columns:138px minmax(0,1fr);
  align-items:center;
  gap:17px;
  padding:16px;
}

.pie {
  position:relative;
  width:132px;
  height:132px;
  border-radius:50%;
  box-shadow:inset 0 0 0 1px var(--border);
}

.pie::after {
  content:"";
  position:absolute;
  inset:26px;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:50%;
}

.pie-center {
  position:absolute;
  inset:0;
  z-index:1;
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;
  color:var(--text);
  font-size:18px;
  font-weight:800;
}

.pie-center small {
  font-size:9px;
  color:var(--muted);
  text-transform:uppercase;
  letter-spacing:.06em;
}

.legend {
  display:grid;
  gap:8px;
  color:var(--text);
  font-size:12px;
  min-width:0;
}

.legend-row {
  display:grid;
  grid-template-columns:11px minmax(0,1fr) auto;
  gap:8px;
  align-items:center;
}

.legend-row span:nth-child(2) {
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
}

.legend-row strong {
  color:var(--text);
  font-size:11.5px;
}

.dot {
  width:10px;
  height:10px;
  border-radius:999px;
}

.dot.green { background:var(--green); }
.dot.red { background:var(--red); }
.dot.blue { background:var(--accent); }
.dot.orange { background:var(--amber); }
.dot.purple { background:var(--purple); }
.dot.gray { background:var(--gray); }

.chart-list-card {
  min-height:220px;
  padding:16px;
}

.chart-list-title {
  margin:0 0 14px;
  color:var(--text);
  font-size:13px;
  font-weight:800;
}

.bar-list {
  display:grid;
  gap:10px;
}

.bar-row {
  display:grid;
  grid-template-columns:minmax(95px,130px) minmax(80px,1fr) 36px;
  align-items:center;
  gap:10px;
  font-size:11.5px;
}

.bar-label {
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
  color:var(--muted);
}

.bar-track {
  height:8px;
  border-radius:999px;
  background:var(--surface3);
  overflow:hidden;
}

.bar-fill {
  height:100%;
  min-width:2px;
  border-radius:999px;
  background:linear-gradient(90deg,var(--accent),var(--navy));
}

.bar-count {
  text-align:right;
  font-weight:800;
}

.toolbar {
  display:flex;
  gap:9px;
  flex-wrap:wrap;
  align-items:flex-start;
  padding:14px;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  box-shadow:var(--shadow);
  margin-bottom:13px;
}

input[type="search"] {
  min-height:38px;
  min-width:340px;
  flex:1 1 340px;
  padding:8px 10px;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius-sm);
  color:var(--text);
  font-size:12.5px;
}

input[type="search"]::placeholder {
  color:var(--muted);
}

.filter-menu {
  position:relative;
}

.filter-trigger {
  min-height:38px;
  min-width:168px;
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:9px;
  padding:8px 10px;
  border:1px solid var(--border);
  border-radius:var(--radius-sm);
  background:var(--surface);
  color:var(--text);
  cursor:pointer;
  font-size:12.5px;
  font-weight:650;
  transition:.15s ease;
}

.filter-trigger:hover,
.filter-trigger[aria-expanded="true"] {
  background:var(--surface2);
}

.filter-label {
  white-space:nowrap;
}

.filter-summary {
  max-width:126px;
  overflow:hidden;
  color:var(--muted);
  font-size:10.5px;
  font-weight:700;
  text-overflow:ellipsis;
  white-space:nowrap;
}

.filter-chevron {
  color:var(--muted);
  font-size:10px;
  transition:transform .15s ease;
}

.filter-trigger[aria-expanded="true"] .filter-chevron {
  transform:rotate(180deg);
}

.filter-popover {
  position:absolute;
  top:calc(100% + 7px);
  right:0;
  z-index:40;
  width:310px;
  max-height:410px;
  display:none;
  overflow:hidden;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  box-shadow:var(--shadow-hover);
}

.filter-popover.open {
  display:block;
}

.filter-popover-head {
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  padding:11px 12px;
  background:var(--surface2);
  border-bottom:1px solid var(--border);
}

.filter-popover-head strong {
  font-size:10.5px;
  text-transform:uppercase;
  letter-spacing:.055em;
}

.filter-link {
  padding:0;
  border:0;
  background:transparent;
  color:var(--accent);
  cursor:pointer;
  font-size:10.5px;
  font-weight:800;
}

.filter-options {
  max-height:340px;
  overflow:auto;
  padding:7px;
}

.check-option {
  display:grid;
  grid-template-columns:18px minmax(0,1fr) auto;
  gap:9px;
  align-items:center;
  min-height:36px;
  padding:7px 8px;
  border-radius:8px;
  cursor:pointer;
}

.check-option:hover {
  background:var(--surface2);
}

.check-option input {
  width:16px;
  height:16px;
  margin:0;
  accent-color:var(--accent);
  cursor:pointer;
}

.check-label {
  min-width:0;
  overflow:hidden;
  color:var(--text);
  font-size:11.5px;
  text-overflow:ellipsis;
  white-space:nowrap;
}

.check-count {
  min-width:28px;
  padding:2px 6px;
  border-radius:999px;
  background:var(--surface3);
  color:var(--muted);
  font-size:9.5px;
  font-weight:800;
  text-align:center;
}

.active-filter-count {
  min-width:19px;
  height:19px;
  max-width:none;
  display:inline-grid;
  place-items:center;
  padding:0 6px;
  border-radius:999px;
  background:var(--accent);
  color:#fff;
  font-size:9.5px;
  font-weight:900;
}

.result-count {
  margin-left:auto;
  min-height:38px;
  display:flex;
  align-items:center;
  padding:0 6px;
  color:var(--muted);
  font-size:11.5px;
}

.table-wrap {
  overflow:auto;
  max-height:760px;
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:var(--radius);
  box-shadow:var(--shadow);
}

table {
  width:100%;
  border-collapse:collapse;
  font-size:12.5px;
  white-space:nowrap;
}

th {
  position:sticky;
  top:0;
  z-index:2;
  padding:11px 12px;
  background:var(--surface2);
  border-bottom:1px solid var(--border);
  color:var(--muted);
  text-align:left;
  font-size:10px;
  font-weight:800;
  text-transform:uppercase;
  letter-spacing:.045em;
  cursor:pointer;
  user-select:none;
}

th[data-sort]::after {
  content:" ↕";
  opacity:.45;
}

td {
  padding:10px 12px;
  border-bottom:1px solid var(--border);
  vertical-align:top;
}

tbody tr.data-row:hover td {
  background:var(--surface2);
}

tbody tr:last-child td {
  border-bottom:0;
}

.name-cell {
  white-space:normal;
  min-width:340px;
  max-width:720px;
}

.test-name {
  font-weight:700;
  color:var(--text);
}

.test-id {
  margin-top:3px;
  color:var(--muted);
  font-size:10.5px;
}

/* =========================================================
   GENERIC PILLS
   Service and Severity keep the existing visual style
   ========================================================= */

.pill {
  display:inline-flex;
  align-items:center;
  padding:3px 8px;
  border:1px solid var(--border);
  border-radius:999px;
  background:var(--surface3);
  color:var(--muted);
  font-size:10.5px;
  font-weight:700;
}

.pill.good {
  background:var(--green-soft);
  border-color:transparent;
  color:var(--green) !important;
}

.pill.bad {
  background:var(--red-soft);
  border-color:transparent;
  color:var(--red) !important;
}

.pill.warn {
  background:var(--amber-soft);
  border-color:transparent;
  color:var(--amber) !important;
}

.pill.investigate {
  background:var(--purple-soft);
  border-color:transparent;
  color:var(--purple) !important;
}

.pill.neutral {
  background:var(--gray-soft);
  border-color:transparent;
  color:var(--gray) !important;
}

.pill.critical {
  background:var(--red-soft);
  border-color:transparent;
  color:var(--red) !important;
}

.pill.high {
  background:var(--amber-soft);
  border-color:transparent;
  color:var(--amber) !important;
}

.pill.medium {
  background:var(--purple-soft);
  border-color:transparent;
  color:var(--purple) !important;
}

.pill.low {
  background:var(--green-soft);
  border-color:transparent;
  color:var(--green) !important;
}


/* =========================================================
   STATUS BADGES
   STYLE ONLY
   Compact result badge matching the supplied screenshot
   ========================================================= */

.data-row td:first-child {
  vertical-align:middle;
}

.data-row td:first-child .pill {
  display:inline-flex;
  align-items:center;
  justify-content:center;
  gap:6px;

  min-height:24px;
  padding:5px 9px;

  border:0;
  border-radius:8px;

  font-size:10.5px;
  font-weight:700;
  line-height:1;
  letter-spacing:0;

  box-shadow:none;
}

/* Small round status dot */
.data-row td:first-child .pill::before {
  content:"";
  display:block;
  width:6px;
  height:6px;
  min-width:6px;
  flex:0 0 6px;
  border-radius:999px;
  background:currentColor;
  opacity:.78;
}

/* FAILED + ERROR - screenshot style */
.data-row td:first-child .pill.bad {
  background:#fde9ec;
  color:#d9485f !important;
}

/* PASSED */
.data-row td:first-child .pill.good {
  background:#e7f7ef;
  color:#27966a !important;
}

/* SKIPPED */
.data-row td:first-child .pill.warn {
  background:#fff4dc;
  color:#c48417 !important;
}

/* INVESTIGATE */
.data-row td:first-child .pill.investigate {
  background:#f3eafb;
  color:#8c55b6 !important;
}

/* NOT RUN */
.data-row td:first-child .pill.neutral {
  background:#eef1f5;
  color:#667085 !important;
}

/* Dark-mode equivalents */
[data-theme="dark"] .data-row td:first-child .pill.bad {
  background:rgba(217,72,95,.16);
  color:#ff8497 !important;
}

[data-theme="dark"] .data-row td:first-child .pill.good {
  background:rgba(39,150,106,.17);
  color:#67d6a6 !important;
}

[data-theme="dark"] .data-row td:first-child .pill.warn {
  background:rgba(196,132,23,.18);
  color:#f2bf5b !important;
}

[data-theme="dark"] .data-row td:first-child .pill.investigate {
  background:rgba(140,85,182,.19);
  color:#d0a4f0 !important;
}

[data-theme="dark"] .data-row td:first-child .pill.neutral {
  background:rgba(102,112,133,.22);
  color:#b9c0ce !important;
}


/* =========================================================
   DETAIL ROW
   ========================================================= */

.detail-row {
  display:none;
}

.detail-row.open {
  display:table-row;
}

.detail-row td {
  padding:0;
  background:var(--surface2);
  white-space:normal;
}

.detail-panel {
  padding:18px 20px 20px;
  display:grid;
  grid-template-columns:repeat(2,minmax(0,1fr));
  gap:16px;
  border-bottom:1px solid var(--border);
}

.detail-block {
  min-width:0;
}

.detail-block.full {
  grid-column:1 / -1;
}

.detail-heading {
  margin-bottom:7px;
  color:var(--muted);
  font-size:10px;
  font-weight:800;
  text-transform:uppercase;
  letter-spacing:.055em;
}

.detail-text {
  margin:0;
  padding:14px;
  min-height:58px;
  max-height:460px;
  overflow:auto;
  border:1px solid var(--border);
  border-radius:var(--radius-sm);
  background:var(--surface);
  color:var(--text);
  font:12px/1.6 var(--font);
  overflow-wrap:anywhere;
}

.detail-text p {
  margin:0 0 10px;
}

.detail-text p:last-child {
  margin-bottom:0;
}

.detail-text ul,
.detail-text ol {
  margin:8px 0 10px;
  padding-left:22px;
}

.detail-text li {
  margin:6px 0;
  padding-left:2px;
}

.detail-text a {
  color:var(--accent);
  font-weight:650;
  text-decoration:none;
  overflow-wrap:anywhere;
}

.detail-text a:hover {
  text-decoration:underline;
}

.detail-text .detail-title {
  margin:12px 0 7px;
  color:var(--text);
  font-size:12px;
  font-weight:800;
}

.detail-text .detail-title:first-child {
  margin-top:0;
}

.detail-text .result-callout {
  display:flex;
  align-items:flex-start;
  gap:9px;
  margin:0 0 12px;
  padding:10px 12px;
  border-radius:var(--radius-sm);
  border:1px solid var(--border);
  background:var(--surface2);
}

.detail-text .result-callout.good {
  background:var(--green-soft);
  border-color:transparent;
  color:var(--green) !important;
}

.detail-text .result-callout.bad {
  background:var(--red-soft);
  border-color:transparent;
  color:var(--red) !important;
}

.detail-text .result-callout.warn {
  background:var(--amber-soft);
  border-color:transparent;
  color:var(--amber) !important;
}

.detail-text .result-icon {
  flex:0 0 auto;
  font-weight:900;
}

.detail-object {
  display:grid;
  gap:10px;
}

.detail-object-row {
  padding:10px 12px;
  border:1px solid var(--border);
  border-radius:var(--radius-sm);
  background:var(--surface2);
}

.detail-object-key {
  margin-bottom:5px;
  color:var(--muted);
  font-size:9.5px;
  font-weight:800;
  text-transform:uppercase;
  letter-spacing:.055em;
}

.metadata-grid {
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
  gap:9px;
}

.metadata-item {
  min-width:0;
  padding:10px 12px;
  border:1px solid var(--border);
  border-radius:var(--radius-sm);
  background:var(--surface2);
}

.metadata-label {
  display:block;
  margin-bottom:3px;
  color:var(--muted);
  font-size:9px;
  font-weight:800;
  text-transform:uppercase;
  letter-spacing:.055em;
}

.metadata-value {
  color:var(--text);
  font-size:11.5px;
  overflow-wrap:anywhere;
}

.metadata-value a {
  color:var(--accent);
  font-weight:700;
  text-decoration:none;
}

.metadata-value a:hover {
  text-decoration:underline;
}

.empty-state {
  display:none;
  padding:46px 20px;
  color:var(--muted);
  text-align:center;
}

footer {
  max-width:1800px;
  margin:0 auto;
  padding:0 32px 32px;
  color:var(--muted);
  font-size:11.5px;
}

@media (max-width:1180px) {
  .mini-grid {
    grid-template-columns:1fr;
  }
}

@media (max-width:820px) {
  .topbar {
    align-items:flex-start;
    padding:12px 16px;
    flex-wrap:wrap;
  }

  .topbar-actions {
    width:100%;
    margin-left:54px;
    flex-wrap:wrap;
  }

  .generated {
    margin-left:auto;
  }

  .layout {
    padding:20px 16px 30px;
  }

  .detail-panel {
    grid-template-columns:1fr;
  }

  .detail-block.full {
    grid-column:auto;
  }

  input[type="search"] {
    min-width:100%;
  }

  .filter-menu {
    flex:1 1 calc(50% - 6px);
  }

  .filter-trigger {
    width:100%;
    min-width:0;
  }

  .filter-popover {
    left:0;
    right:auto;
    width:min(330px,calc(100vw - 32px));
  }
}

@media print {
  header {
    position:static;
  }

  .topbar-actions .btn,
  .toolbar,
  .details-button {
    display:none !important;
  }

  .table-wrap {
    max-height:none;
    overflow:visible;
  }

  .table-wrap,
  .grid > .card,
  .section .card {
    box-shadow:none;
  }

  .detail-row {
    display:none !important;
  }
}
</style>
</head>

<body>

<header>
  <div class="topbar">
    <div class="brand-left">
      <div class="logo-fallback" aria-hidden="true">M</div>
      <div>
        <h1>ITI365 - Security Assessment</h1>
        <div class="subtitle" id="tenantSubtitle"></div>
      </div>
    </div>

    <div class="topbar-actions">
      <a id="originalReportButton" class="btn" href="#">Original report</a>
      <button class="btn" type="button" id="exportButton">Export CSV</button>
      <button class="btn" type="button" onclick="window.print()">Print</button>
      <button class="btn" type="button" id="themeButton">Dark mode</button>
    </div>

    <div class="generated" id="generatedBlock"></div>
  </div>
</header>

<main class="layout">

  <section class="grid" aria-label="Assessment summary">

    <article class="card">
      <div class="card-title">Security score</div>
      <div class="card-value info" id="scoreValue">0%</div>
      <div class="card-note">Passed checks among assessed results</div>
    </article>

    <article class="card card-good">
      <div class="card-title">Passed</div>
      <div class="card-value good" id="passedValue">0</div>
      <div class="card-note">Checks completed successfully</div>
    </article>

    <article class="card card-bad">
      <div class="card-title">Failed</div>
      <div class="card-value bad" id="failedValue">0</div>
      <div class="card-note">Configuration changes recommended</div>
    </article>

    <article class="card card-investigate">
      <div class="card-title">Investigate</div>
      <div class="card-value investigate" id="investigateValue">0</div>
      <div class="card-note">Manual validation required</div>
    </article>

    <article class="card card-warn">
      <div class="card-title">Skipped / not run</div>
      <div class="card-value warn" id="skippedValue">0</div>
      <div class="card-note">Licensing, connection or applicability</div>
    </article>

    <article class="card">
      <div class="card-title">Total tests</div>
      <div class="card-value" id="totalValue">0</div>
      <div class="card-note" id="durationNote">Complete Maester assessment</div>
    </article>

  </section>

  <section class="section">
    <h2>Assessment overview</h2>

    <div class="mini-grid">

      <article class="card chart-card">
        <div class="pie" id="statusPie">
          <div class="pie-center">
            <span id="statusPieValue">0</span>
            <small>tests</small>
          </div>
        </div>

        <div>
          <h3 class="chart-list-title">Result distribution</h3>
          <div class="legend" id="statusLegend"></div>
        </div>
      </article>

      <article class="card chart-list-card">
        <h3 class="chart-list-title">Severity distribution</h3>
        <div class="bar-list" id="severityBars"></div>
      </article>

      <article class="card chart-list-card">
        <h3 class="chart-list-title">Tests by service</h3>
        <div class="bar-list" id="serviceBars"></div>
      </article>

    </div>
  </section>

  <section class="section">
    <h2>Detailed test analysis</h2>

    <div class="toolbar">

      <input
        id="searchInput"
        type="search"
        placeholder="Search test name, ID, service, description or result…"
        autocomplete="off"
      >

      <div class="filter-menu">
        <button
          class="filter-trigger"
          type="button"
          data-filter-trigger="statuses"
          aria-expanded="false"
          aria-controls="statusesMenu"
        >
          <span class="filter-label">Status</span>
          <span class="filter-summary" id="statusesSummary">All</span>
          <span class="filter-chevron">▼</span>
        </button>

        <div class="filter-popover" id="statusesMenu">
          <div class="filter-popover-head">
            <strong>Status</strong>
            <button class="filter-link" type="button" data-clear-filter="statuses">Clear all</button>
          </div>

          <div class="filter-options" id="statusesOptions"></div>
        </div>
      </div>

      <div class="filter-menu">
        <button
          class="filter-trigger"
          type="button"
          data-filter-trigger="severities"
          aria-expanded="false"
          aria-controls="severitiesMenu"
        >
          <span class="filter-label">Severity</span>
          <span class="filter-summary" id="severitiesSummary">All</span>
          <span class="filter-chevron">▼</span>
        </button>

        <div class="filter-popover" id="severitiesMenu">
          <div class="filter-popover-head">
            <strong>Severity</strong>
            <button class="filter-link" type="button" data-clear-filter="severities">Clear all</button>
          </div>

          <div class="filter-options" id="severitiesOptions"></div>
        </div>
      </div>

      <div class="filter-menu">
        <button
          class="filter-trigger"
          type="button"
          data-filter-trigger="services"
          aria-expanded="false"
          aria-controls="servicesMenu"
        >
          <span class="filter-label">Service</span>
          <span class="filter-summary" id="servicesSummary">All</span>
          <span class="filter-chevron">▼</span>
        </button>

        <div class="filter-popover" id="servicesMenu">
          <div class="filter-popover-head">
            <strong>Service</strong>
            <button class="filter-link" type="button" data-clear-filter="services">Clear all</button>
          </div>

          <div class="filter-options" id="servicesOptions"></div>
        </div>
      </div>

      <button class="btn" type="button" id="clearButton">Clear filters</button>

      <div class="result-count">
        <strong id="visibleCount">0</strong>&nbsp;visible
      </div>

    </div>

    <div class="table-wrap">
      <table id="resultsTable">
        <thead>
          <tr>
            <th data-sort="Status">Status</th>
            <th data-sort="Name">Test</th>
            <th data-sort="Service">Service</th>
            <th data-sort="Severity">Severity</th>
            <th data-sort="Duration">Duration</th>
            <th>Analysis</th>
          </tr>
        </thead>

        <tbody id="resultsBody"></tbody>
      </table>

      <div class="empty-state" id="emptyState">
        No tests match the current filters.
      </div>
    </div>

  </section>

</main>

<footer>
  Generated from ITI365 execution as the preserved original report.
  This dashboard is self-contained and can be archived or shared as a single HTML file.
</footer>

<script>
const report = __DATA_JSON__;

let sortState = {
  key: 'Status',
  direction: 'asc'
};

let visibleTests = [...report.tests];

const filterState = {
  statuses: new Set(),
  severities: new Set(),
  services: new Set()
};

const statusColors = {
  Passed: 'var(--green)',
  Failed: 'var(--red)',
  Investigate: 'var(--purple)',
  Skipped: 'var(--amber)',
  NotRun: 'var(--gray)',
  Error: 'var(--red)'
};

const statusDotClasses = {
  Passed: 'green',
  Failed: 'red',
  Investigate: 'purple',
  Skipped: 'orange',
  NotRun: 'gray',
  Error: 'red'
};

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function statusClass(status) {
  if (status === 'Passed') return 'good';
  if (status === 'Failed' || status === 'Error') return 'bad';
  if (status === 'Investigate') return 'investigate';
  if (status === 'Skipped') return 'warn';

  return 'neutral';
}

function severityClass(severity) {
  const value = String(severity || '').toLowerCase();

  if (['critical','high','medium','low'].includes(value)) {
    return value;
  }

  return 'neutral';
}

function setSummary() {
  document.getElementById('tenantSubtitle').textContent = [
    report.tenant.name,
    report.tenant.id ? `Tenant ID: ${report.tenant.id}` : '',
    report.tenant.moduleVersion ? `Maester ${report.tenant.moduleVersion}` : ''
  ].filter(Boolean).join(' • ');

  document.getElementById('generatedBlock').innerHTML =
    `<strong>Generated</strong><br>${escapeHtml(report.tenant.generatedAt)}`;

  const originalButton = document.getElementById('originalReportButton');
  originalButton.href = encodeURI(report.tenant.originalReport);
  originalButton.target = '_blank';
  originalButton.rel = 'noopener';

  document.getElementById('scoreValue').textContent =
    `${report.summary.score}%`;

  document.getElementById('passedValue').textContent =
    report.summary.passed;

  document.getElementById('failedValue').textContent =
    report.summary.failed;

  document.getElementById('investigateValue').textContent =
    report.summary.investigate;

  document.getElementById('skippedValue').textContent =
    report.summary.skipped + report.summary.notRun;

  document.getElementById('totalValue').textContent =
    report.summary.total;

  document.getElementById('durationNote').textContent =
    report.tenant.runDuration
      ? `Completed in ${report.tenant.runDuration}`
      : 'Complete Maester assessment';
}

function createConicGradient(counts) {
  const entries = Object.entries(counts)
    .filter(([, count]) => Number(count) > 0);

  const total = entries.reduce(
    (sum, [, count]) => sum + Number(count),
    0
  );

  if (!total) {
    return 'var(--surface3)';
  }

  let current = 0;

  const slices = entries.map(([name, count]) => {
    const start = (current / total) * 360;
    current += Number(count);
    const end = (current / total) * 360;

    return `${statusColors[name] || 'var(--accent)'} ${start}deg ${end}deg`;
  });

  return `conic-gradient(${slices.join(',')})`;
}

function renderCharts() {
  const statusPie = document.getElementById('statusPie');

  statusPie.style.background =
    createConicGradient(report.statusCounts);

  document.getElementById('statusPieValue').textContent =
    report.summary.total;

  const statusLegend =
    document.getElementById('statusLegend');

  statusLegend.innerHTML =
    Object.entries(report.statusCounts).map(([name, count]) => `
      <div class="legend-row">
        <span class="dot ${statusDotClasses[name] || 'blue'}"></span>
        <span>${escapeHtml(name)}</span>
        <strong>${count}</strong>
      </div>
    `).join('');

  renderBars('severityBars', report.severityCounts, 7);
  renderBars('serviceBars', report.serviceCounts, 8);
}

function renderBars(elementId, counts, limit) {
  const entries =
    Object.entries(counts).slice(0, limit);

  const max =
    Math.max(1, ...entries.map(([, count]) => Number(count)));

  document.getElementById(elementId).innerHTML =
    entries.map(([name, count]) => `
      <div class="bar-row" title="${escapeHtml(name)}: ${count}">
        <div class="bar-label">${escapeHtml(name)}</div>

        <div class="bar-track">
          <div
            class="bar-fill"
            style="width:${Math.max(2, (Number(count) / max) * 100)}%">
          </div>
        </div>

        <div class="bar-count">${count}</div>
      </div>
    `).join('');
}

function uniqueValues(property) {
  return [
    ...new Set(
      report.tests
        .map(test => test[property])
        .filter(Boolean)
    )
  ].sort((a, b) =>
    String(a).localeCompare(String(b))
  );
}

function countValues(values) {
  return values.reduce((counts, value) => {
    counts.set(value, (counts.get(value) || 0) + 1);
    return counts;
  }, new Map());
}

const filterDefinitions = {
  statuses: {
    property: 'Status',
    values: uniqueValues('Status'),
    counts: countValues(
      report.tests.map(test => test.Status).filter(Boolean)
    ),
    optionsId: 'statusesOptions',
    summaryId: 'statusesSummary'
  },

  severities: {
    property: 'Severity',
    values: uniqueValues('Severity'),
    counts: countValues(
      report.tests.map(test => test.Severity).filter(Boolean)
    ),
    optionsId: 'severitiesOptions',
    summaryId: 'severitiesSummary'
  },

  services: {
    property: 'Service',
    values: uniqueValues('Service'),
    counts: countValues(
      report.tests.map(test => test.Service).filter(Boolean)
    ),
    optionsId: 'servicesOptions',
    summaryId: 'servicesSummary'
  }
};

function renderFilterOptions(filterName) {
  const definition =
    filterDefinitions[filterName];

  const selected =
    filterState[filterName];

  const container =
    document.getElementById(definition.optionsId);

  container.innerHTML =
    definition.values.map((value, index) => {
      const id = `${filterName}-${index}`;

      return `
        <label
          class="check-option"
          for="${id}"
          title="${escapeHtml(value)}"
        >
          <input
            id="${id}"
            type="checkbox"
            data-filter-name="${filterName}"
            value="${escapeHtml(value)}"
            ${selected.has(value) ? 'checked' : ''}
          >

          <span class="check-label">
            ${escapeHtml(value)}
          </span>

          <span class="check-count">
            ${definition.counts.get(value) || 0}
          </span>
        </label>
      `;
    }).join('');

  updateFilterSummary(filterName);
}

function updateFilterSummary(filterName) {
  const selected =
    filterState[filterName];

  const summary =
    document.getElementById(
      filterDefinitions[filterName].summaryId
    );

  if (!selected.size) {
    summary.textContent = 'All';
    summary.classList.remove('active-filter-count');
    return;
  }

  summary.textContent =
    selected.size === 1
      ? [...selected][0]
      : `${selected.size} selected`;

  summary.classList.toggle(
    'active-filter-count',
    selected.size > 1
  );
}

function closeFilterMenus(exceptName = '') {
  document
    .querySelectorAll('[data-filter-trigger]')
    .forEach(trigger => {

      const filterName =
        trigger.dataset.filterTrigger;

      const menu =
        document.getElementById(`${filterName}Menu`);

      const keepOpen =
        filterName === exceptName;

      trigger.setAttribute(
        'aria-expanded',
        keepOpen ? 'true' : 'false'
      );

      menu.classList.toggle(
        'open',
        keepOpen
      );
    });
}

function safeUrl(value) {
  const url =
    String(value || '').trim();

  return /^https?:\/\//i.test(url)
    ? url
    : '';
}

function createLinkToken(label, url, tokens) {
  const safe = safeUrl(url);

  if (!safe) {
    return label;
  }

  const token =
    `@@MAESTER_LINK_${tokens.length}@@`;

  tokens.push(
    `<a href="${escapeHtml(safe)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label || safe)}</a>`
  );

  return token;
}

function linkifyText(value) {
  let text =
    String(value ?? '');

  const tokens = [];

  text = text.replace(
    /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/gi,
    (_, label, url) =>
      createLinkToken(label, url, tokens)
  );

  text = text.replace(
    /https?:\/\/[^\s<>"')\]]+/gi,
    url =>
      createLinkToken(url, url, tokens)
  );

  let encoded =
    escapeHtml(text);

  tokens.forEach((link, index) => {
    encoded = encoded.replace(
      `@@MAESTER_LINK_${index}@@`,
      link
    );
  });

  return encoded;
}

function tryParseDetailJson(value) {
  if (typeof value !== 'string') {
    return value;
  }

  const text =
    value.trim();

  if (
    !text ||
    (!text.startsWith('{') &&
     !text.startsWith('['))
  ) {
    return value;
  }

  try {
    return JSON.parse(text);
  }
  catch {
    return value;
  }
}

function prettifyKey(value) {
  return String(value || '')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replaceAll('_', ' ')
    .replace(/\b\w/g, char => char.toUpperCase());
}

function removeExcludedMarkdownSections(value) {
  const text =
    String(value ?? '')
      .replace(/\r\n?/g, '\n')
      .replace(/\\n/g, '\n')
      .replace(/\\t/g, '  ');

  const output = [];
  let skipSection = false;
  let skipHeadingLevel = 0;

  for (const rawLine of text.split('\n')) {

    const heading =
      rawLine.match(/^\s*(#{1,6})\s+(.+?)\s*$/);

    if (skipSection) {
      if (
        heading &&
        heading[1].length <= skipHeadingLevel
      ) {
        skipSection = false;
        skipHeadingLevel = 0;
      }
      else {
        continue;
      }
    }

    if (heading) {
      const headingText =
        heading[2]
          .trim()
          .replace(/^[*_`\s]+|[*_`\s]+$/g, '');

      if (
        /^(?:Remediation\s+action|Related\s+links)\s*:?\s*$/i
          .test(headingText)
      ) {
        skipSection = true;
        skipHeadingLevel = heading[1].length;
        continue;
      }
    }

    output.push(rawLine);
  }

  return output
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function formatTextContent(value, tone = '') {
  const text =
    removeExcludedMarkdownSections(value);

  if (!text) {
    return '<p>No information was provided.</p>';
  }

  const lines =
    text.split('\n');

  const output = [];

  let listType = '';

  const closeList = () => {
    if (listType) {
      output.push(`</${listType}>`);
      listType = '';
    }
  };

  for (const originalLine of lines) {
    const line =
      originalLine.trim();

    if (!line) {
      closeList();
      continue;
    }

    const bullet =
      line.match(/^[-*•]\s+(.+)$/);

    const numbered =
      line.match(/^\d+[.)]\s+(.+)$/);

    if (bullet || numbered) {
      const requiredType =
        bullet ? 'ul' : 'ol';

      if (listType !== requiredType) {
        closeList();
        listType = requiredType;
        output.push(`<${listType}>`);
      }

      output.push(
        `<li>${linkifyText((bullet || numbered)[1])}</li>`
      );

      continue;
    }

    closeList();

    const heading =
      line.match(/^#{1,5}\s+(.+)$/);

    if (heading) {
      output.push(
        `<div class="detail-title">${linkifyText(heading[1])}</div>`
      );

      continue;
    }

    const isSuccess =
      /^(well done|passed|success)/i.test(line);

    const isFailure =
      /^(failed|failure|error)/i.test(line);

    if (isSuccess || isFailure) {
      const calloutTone =
        isSuccess ? 'good' : 'bad';

      const icon =
        isSuccess ? '✓' : '!';

      output.push(
        `<div class="result-callout ${calloutTone}"><span class="result-icon">${icon}</span><span>${linkifyText(line)}</span></div>`
      );

      continue;
    }

    output.push(
      `<p>${linkifyText(line)}</p>`
    );
  }

  closeList();

  return output.join('');
}

function formatDetailValue(value, tone = '') {
  const parsed =
    tryParseDetailJson(value);

  if (
    parsed === null ||
    parsed === undefined ||
    parsed === ''
  ) {
    return '<p>No information was provided.</p>';
  }

  if (Array.isArray(parsed)) {
    if (!parsed.length) {
      return '<p>No information was provided.</p>';
    }

    return `<ul>${
      parsed.map(item =>
        `<li>${
          typeof item === 'object'
            ? formatDetailValue(item, tone)
            : linkifyText(item)
        }</li>`
      ).join('')
    }</ul>`;
  }

  if (typeof parsed === 'object') {
    const entries =
      Object.entries(parsed).filter(([, item]) =>
        item !== null &&
        item !== undefined &&
        String(item).trim() !== ''
      );

    if (!entries.length) {
      return '<p>No information was provided.</p>';
    }

    return `
      <div class="detail-object">
        ${
          entries.map(([key, item]) => `
            <div class="detail-object-row">
              <div class="detail-object-key">
                ${escapeHtml(prettifyKey(key))}
              </div>

              <div>
                ${formatDetailValue(item, tone)}
              </div>
            </div>
          `).join('')
        }
      </div>
    `;
  }

  return formatTextContent(parsed, tone);
}

function formatMetadata(test) {
  const items = [
    ['Control ID', test.ControlId],
    ['Service', test.Service],
    ['Category', test.Category],
    ['Severity', test.Severity],
    ['Tags', test.Tags],
    ['Duration', test.Duration],
    ['Source', test.Source],
    ['Documentation', test.HelpUrl]
  ].filter(([, value]) => value);

  if (!items.length) {
    return '<p>No additional metadata.</p>';
  }

  return `
    <div class="metadata-grid">
      ${
        items.map(([label, value]) => {

          const safe =
            label === 'Documentation'
              ? safeUrl(value)
              : '';

          const rendered =
            safe
              ? `<a href="${escapeHtml(safe)}" target="_blank" rel="noopener noreferrer">Open Maester documentation ↗</a>`
              : escapeHtml(value);

          return `
            <div class="metadata-item">
              <span class="metadata-label">
                ${escapeHtml(label)}
              </span>

              <div class="metadata-value">
                ${rendered}
              </div>
            </div>
          `;
        }).join('')
      }
    </div>
  `;
}

function compareValues(a, b, key) {
  const av =
    a[key] ?? '';

  const bv =
    b[key] ?? '';

  if (key === 'Status') {
    const order = {
      Failed: 1,
      Error: 2,
      Investigate: 3,
      Skipped: 4,
      NotRun: 5,
      Passed: 6
    };

    return (order[av] || 99) -
           (order[bv] || 99);
  }

  if (key === 'Severity') {
    const order = {
      Critical: 1,
      High: 2,
      Medium: 3,
      Low: 4,
      Info: 5,
      'Not specified': 6
    };

    return (order[av] || 99) -
           (order[bv] || 99);
  }

  return String(av).localeCompare(
    String(bv),
    undefined,
    {
      numeric: true,
      sensitivity: 'base'
    }
  );
}

function applyFilters() {
  const query =
    document
      .getElementById('searchInput')
      .value
      .trim()
      .toLowerCase();

  visibleTests =
    report.tests.filter(test => {

      if (
        filterState.statuses.size &&
        !filterState.statuses.has(test.Status)
      ) {
        return false;
      }

      if (
        filterState.severities.size &&
        !filterState.severities.has(test.Severity)
      ) {
        return false;
      }

      if (
        filterState.services.size &&
        !filterState.services.has(test.Service)
      ) {
        return false;
      }

      if (query) {
        const haystack = [
          test.ControlId,
          test.Name,
          test.Status,
          test.Severity,
          test.Service,
          test.Category,
          test.Tags,
          test.Description,
          test.Details,
          test.SkipReason,
          test.Error,
          test.Source,
          test.HelpUrl
        ]
        .join(' ')
        .toLowerCase();

        if (!haystack.includes(query)) {
          return false;
        }
      }

      return true;
    });

  visibleTests.sort((a, b) => {
    const result =
      compareValues(a, b, sortState.key);

    return sortState.direction === 'asc'
      ? result
      : -result;
  });

  renderTable();
}

function renderTable() {
  const body =
    document.getElementById('resultsBody');

  body.innerHTML = '';

  visibleTests.forEach(test => {

    const dataRow =
      document.createElement('tr');

    dataRow.className =
      'data-row';

    dataRow.innerHTML = `
      <td>
        <span class="pill ${statusClass(test.Status)}">
          ${escapeHtml(test.Status)}
        </span>
      </td>

      <td class="name-cell">
        <div class="test-name">
          ${escapeHtml(test.Name)}
        </div>

        <div class="test-id">
          ${escapeHtml(test.ControlId || test.Tags || '')}
        </div>
      </td>

      <td>
        <span class="pill">
          ${escapeHtml(test.Service || 'Other')}
        </span>
      </td>

      <td>
        <span class="pill ${severityClass(test.Severity)}">
          ${escapeHtml(test.Severity)}
        </span>
      </td>

      <td>
        ${escapeHtml(test.Duration || '—')}
      </td>

      <td>
        <button
          type="button"
          class="btn details-button"
        >
          View details
        </button>
      </td>
    `;

    const detailRow =
      document.createElement('tr');

    detailRow.className =
      'detail-row';

    detailRow.innerHTML = `
      <td colspan="6">
        <div class="detail-panel">

          <div class="detail-block">
            <div class="detail-heading">
              What is being tested
            </div>

            <div class="detail-text"></div>
          </div>

          <div class="detail-block">
            <div class="detail-heading">
              Readable test result
            </div>

            <div class="detail-text"></div>
          </div>

          <div class="detail-block full">
            <div class="detail-heading">
              Technical metadata
            </div>

            <div class="detail-text"></div>
          </div>

        </div>
      </td>
    `;

    const detailElements =
      detailRow.querySelectorAll('.detail-text');

    detailElements[0].innerHTML =
      formatDetailValue(
        test.Description ||
        'No description was provided by this test.'
      );

    detailElements[1].innerHTML =
      formatDetailValue(
        test.Details ||
        test.SkipReason ||
        test.Error ||
        'No additional result details were provided.',
        statusClass(test.Status)
      );

    detailElements[2].innerHTML =
      formatMetadata(test);

    const button =
      dataRow.querySelector('.details-button');

    button.addEventListener('click', () => {
      const open =
        detailRow.classList.toggle('open');

      button.textContent =
        open
          ? 'Hide details'
          : 'View details';
    });

    body.appendChild(dataRow);
    body.appendChild(detailRow);
  });

  document.getElementById('visibleCount').textContent =
    visibleTests.length;

  document.getElementById('emptyState').style.display =
    visibleTests.length
      ? 'none'
      : 'block';

  document.getElementById('resultsTable').style.display =
    visibleTests.length
      ? 'table'
      : 'none';
}

function csvEscape(value) {
  const text =
    String(value ?? '').replaceAll('"', '""');

  return `"${text}"`;
}

function exportCsv() {
  const columns = [
    'ControlId',
    'Name',
    'Status',
    'Severity',
    'Service',
    'Category',
    'Duration',
    'Tags',
    'Description',
    'Details',
    'SkipReason',
    'Error',
    'Source',
    'HelpUrl'
  ];

  const lines = [
    columns.map(csvEscape).join(',')
  ];

  for (const test of visibleTests) {
    lines.push(
      columns
        .map(column => csvEscape(test[column]))
        .join(',')
    );
  }

  const blob =
    new Blob(
      ['\ufeff' + lines.join('\r\n')],
      {
        type:'text/csv;charset=utf-8'
      }
    );

  const url =
    URL.createObjectURL(blob);

  const link =
    document.createElement('a');

  link.href = url;

  link.download =
    `Maester-Dashboard-${
      report.tenant.generatedAt
        .replaceAll(':','-')
        .replace(' ','_')
    }.csv`;

  document.body.appendChild(link);

  link.click();
  link.remove();

  URL.revokeObjectURL(url);
}

function toggleTheme() {
  const root =
    document.documentElement;

  const next =
    root.getAttribute('data-theme') === 'dark'
      ? 'light'
      : 'dark';

  root.setAttribute(
    'data-theme',
    next
  );

  localStorage.setItem(
    'maester-dashboard-theme',
    next
  );

  document.getElementById('themeButton').textContent =
    next === 'dark'
      ? 'Light mode'
      : 'Dark mode';
}

function initialize() {
  const savedTheme =
    localStorage.getItem(
      'maester-dashboard-theme'
    );

  if (
    savedTheme === 'dark' ||
    savedTheme === 'light'
  ) {
    document.documentElement.setAttribute(
      'data-theme',
      savedTheme
    );
  }
  else if (
    window.matchMedia &&
    window.matchMedia(
      '(prefers-color-scheme: dark)'
    ).matches
  ) {
    document.documentElement.setAttribute(
      'data-theme',
      'dark'
    );
  }

  document.getElementById('themeButton').textContent =
    document.documentElement.getAttribute('data-theme') === 'dark'
      ? 'Light mode'
      : 'Dark mode';

  setSummary();
  renderCharts();

  Object
    .keys(filterDefinitions)
    .forEach(renderFilterOptions);

  document
    .getElementById('searchInput')
    .addEventListener(
      'input',
      applyFilters
    );

  document
    .querySelector('.toolbar')
    .addEventListener('change', event => {

      const checkbox =
        event.target.closest(
          'input[type="checkbox"][data-filter-name]'
        );

      if (!checkbox) {
        return;
      }

      const selected =
        filterState[
          checkbox.dataset.filterName
        ];

      if (checkbox.checked) {
        selected.add(checkbox.value);
      }
      else {
        selected.delete(checkbox.value);
      }

      updateFilterSummary(
        checkbox.dataset.filterName
      );

      applyFilters();
    });

  document
    .querySelectorAll('[data-filter-trigger]')
    .forEach(trigger => {

      trigger.addEventListener(
        'click',
        event => {

          event.stopPropagation();

          const filterName =
            trigger.dataset.filterTrigger;

          const willOpen =
            trigger.getAttribute(
              'aria-expanded'
            ) !== 'true';

          closeFilterMenus(
            willOpen
              ? filterName
              : ''
          );
        }
      );
    });

  document
    .querySelectorAll('.filter-popover')
    .forEach(menu => {

      menu.addEventListener(
        'click',
        event => event.stopPropagation()
      );
    });

  document
    .querySelectorAll('[data-clear-filter]')
    .forEach(button => {

      button.addEventListener(
        'click',
        () => {

          const filterName =
            button.dataset.clearFilter;

          filterState[
            filterName
          ].clear();

          renderFilterOptions(
            filterName
          );

          applyFilters();
        }
      );
    });

  document.addEventListener(
    'click',
    () => closeFilterMenus()
  );

  document.addEventListener(
    'keydown',
    event => {
      if (event.key === 'Escape') {
        closeFilterMenus();
      }
    }
  );

  document
    .getElementById('exportButton')
    .addEventListener(
      'click',
      exportCsv
    );

  document
    .getElementById('themeButton')
    .addEventListener(
      'click',
      toggleTheme
    );

  document
    .getElementById('clearButton')
    .addEventListener(
      'click',
      () => {

        document
          .getElementById('searchInput')
          .value = '';

        Object
          .keys(filterState)
          .forEach(filterName =>
            filterState[filterName].clear()
          );

        Object
          .keys(filterDefinitions)
          .forEach(renderFilterOptions);

        closeFilterMenus();
        applyFilters();
      }
    );

  document
    .querySelectorAll('th[data-sort]')
    .forEach(header => {

      header.addEventListener(
        'click',
        () => {

          const key =
            header.dataset.sort;

          if (sortState.key === key) {
            sortState.direction =
              sortState.direction === 'asc'
                ? 'desc'
                : 'asc';
          }
          else {
            sortState.key = key;
            sortState.direction = 'asc';
          }

          applyFilters();
        }
      );
    });

  applyFilters();
}

initialize();
</script>

</body>
</html>
'@

    $html = $htmlTemplate.Replace('__DATA_JSON__', $dataJson)
    $html | Out-File -FilePath $OutputPath -Encoding utf8
}

$useExistingRun = -not [string]::IsNullOrWhiteSpace($ExistingRunFolder)

if ($useExistingRun) {
    Write-Step 'Using an existing completed Maester run'
    $runFolder = (Resolve-Path -Path $ExistingRunFolder -ErrorAction Stop).Path
    $maesterResults = $null

    Write-Host "Existing run folder: $runFolder" -ForegroundColor Green
}
else {
    Write-Step 'Checking the Maester module'

    $maesterModule =
        Get-LatestAvailableModule -Name 'Maester'

    if ($null -eq $maesterModule) {
        throw @"
    The Maester module is not installed.
    Install it with:

        Install-Module Maester -Scope CurrentUser -Force

    Then install the tests with:

        New-Item -ItemType Directory -Path "$HOME\maester-tests" -Force
        Set-Location "$HOME\maester-tests"
        Install-MaesterTests
"@
    }

    Write-Host "Using Maester module version $($maesterModule.Version)" -ForegroundColor Green

    $resolvedTestsPath =
        (Resolve-Path -Path $TestsPath).Path

    $testFiles = @(
        Get-ChildItem `
            -Path $resolvedTestsPath `
            -Filter '*.Tests.ps1' `
            -File `
            -Recurse `
            -ErrorAction Stop
    )

    if ($testFiles.Count -eq 0) {
        throw "No *.Tests.ps1 files were found under '$resolvedTestsPath'. Run Install-MaesterTests in that folder first."
    }

    $resolvedOutputRoot =
        [System.IO.Path]::GetFullPath($OutputRoot)

    $runStamp =
        Get-Date -Format 'yyyy-MM-dd_HHmmss'

    $runFolder =
        Join-Path `
            -Path $resolvedOutputRoot `
            -ChildPath $runStamp

    New-Item `
        -Path $runFolder `
        -ItemType Directory `
        -Force |
        Out-Null

    if (-not $SkipConnect) {

        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            $TenantId =
                Read-Host 'Enter the target Microsoft Entra tenant ID'
        }

        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            $UserPrincipalName =
                Read-Host 'Enter the exact administrator UPN to use for all new connections'
        }

        $parsedTenantId =
            [guid]::Empty

        if (-not [guid]::TryParse(
            $TenantId,
            [ref]$parsedTenantId
        )) {
            throw "TenantId '$TenantId' is not a valid GUID."
        }

        $TenantId =
            $parsedTenantId.ToString()

        if ($UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+$') {
            throw "UserPrincipalName '$UserPrincipalName' is not a valid UPN."
        }

        if (-not $SkipForcedDisconnect) {
            Disconnect-AllMaesterServices
        }
        else {
            Write-Warning 'Forced disconnect was skipped because -SkipForcedDisconnect was specified.'
        }

        Write-Step 'Connecting to Maester data sources using conflict-safe ordering'

        $connectionState =
            Connect-MaesterServicesSafely `
                -UPN $UserPrincipalName `
                -Tenant $TenantId `
                -SkipAzure:$SkipAzureConnection `
                -RequireAll:$RequireAllServices
    }
    else {
        Import-LatestModule `
            -Name 'Microsoft.Graph.Authentication' `
            -Required |
            Out-Null

        Import-LatestModule `
            -Name 'Maester' `
            -Required |
            Out-Null

        $graphContext =
            Get-MgContext `
                -ErrorAction SilentlyContinue

        if ($null -eq $graphContext) {
            throw '-SkipConnect was specified, but this clean process has no Microsoft Graph context. Remove -SkipConnect or authenticate inside this script.'
        }

        Write-Warning 'Skipping service connections because -SkipConnect was specified.'
    }

    Write-Step "Running $($testFiles.Count) discovered Maester test files"

    $invokeParameters = @{
        Path                 = $resolvedTestsPath
        OutputFolder         = $runFolder
        OutputFolderFileName = 'Maester-Raw'
        PassThru             = $true
        NonInteractive       = $true
        SkipGraphConnect     = $true
        Verbosity            = $Verbosity
    }

    if ($IncludeLongRunning) {
        $invokeParameters.IncludeLongRunning =
            $true
    }

    if ($IncludePreview) {
        $invokeParameters.IncludePreview =
            $true
    }

    $maesterResults =
        Invoke-Maester @invokeParameters
}

$originalHtmlPath =
    Join-Path `
        -Path $runFolder `
        -ChildPath 'Maester-Raw.html'

$jsonPath =
    Join-Path `
        -Path $runFolder `
        -ChildPath 'Maester-Raw.json'

$dashboardPath =
    Join-Path `
        -Path $runFolder `
        -ChildPath 'Maester-Dashboard.html'

$generationLogPath =
    Join-Path `
        -Path $runFolder `
        -ChildPath 'Dashboard-Generation.log'

$waitDeadline =
    (Get-Date).AddSeconds(30)

do {
    $htmlReady =
        Test-Path `
            -Path $originalHtmlPath `
            -PathType Leaf

    $jsonReady =
        Test-Path `
            -Path $jsonPath `
            -PathType Leaf

    if (
        $htmlReady -and
        ($jsonReady -or $null -ne $maesterResults)
    ) {
        break
    }

    Start-Sleep -Milliseconds 500
}
while ((Get-Date) -lt $waitDeadline)

if (-not (
    Test-Path `
        -Path $originalHtmlPath `
        -PathType Leaf
)) {
    throw "Maester completed without creating the expected HTML file: $originalHtmlPath"
}

try {

    if (
        Test-Path `
            -Path $jsonPath `
            -PathType Leaf
    ) {
        Write-Step 'Loading the structured Maester JSON report'

        $reportData =
            Get-Content `
                -Path $jsonPath `
                -Raw `
                -Encoding utf8 |
            ConvertFrom-Json -Depth 100
    }
    elseif ($null -ne $maesterResults) {
        Write-Warning 'Maester-Raw.json was not available. Using the in-memory PassThru result.'

        $reportData =
            $maesterResults
    }
    else {
        throw 'Invoke-Maester returned no result object and no JSON report was available.'
    }

    Write-Step 'Normalizing Maester test results'

    $normalizedTests = @(
        New-NormalizedTestCollection `
            -MaesterResults $reportData
    )

    if ($normalizedTests.Count -eq 0) {
        throw 'The Maester result did not contain any test records. Review Maester-Raw.html and Maester-Raw.json for details.'
    }

    Write-Step 'Creating the modern standalone HTML dashboard'

    New-ModernDashboardHtml `
        -MaesterResults $reportData `
        -Tests $normalizedTests `
        -OriginalHtmlFileName (
            [System.IO.Path]::GetFileName(
                $originalHtmlPath
            )
        ) `
        -OutputPath $dashboardPath

    if (-not (
        Test-Path `
            -Path $dashboardPath `
            -PathType Leaf
    )) {
        throw "Dashboard generation completed without creating the expected file: $dashboardPath"
    }

    @(
        "Dashboard generated successfully: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Source JSON: $jsonPath"
        "Output HTML: $dashboardPath"
        "Normalized tests: $($normalizedTests.Count)"
    ) |
    Set-Content `
        -Path $generationLogPath `
        -Encoding utf8
}
catch {

    $failure =
        $_

    $diagnosticText = @(
        'ITI365 dashboard post-processing failed.'
        "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Run folder: $runFolder"
        "Raw HTML: $originalHtmlPath"
        "Raw JSON: $jsonPath"
        ''
        'Exception:'
        ($failure | Format-List * -Force | Out-String)
        ''
        'Script stack trace:'
        $failure.ScriptStackTrace
    ) -join "`r`n"

    $diagnosticText |
        Set-Content `
            -Path $generationLogPath `
            -Encoding utf8

    $safeMessage =
        [System.Net.WebUtility]::HtmlEncode(
            $failure.Exception.Message
        )

    $safeRawName =
        [System.Net.WebUtility]::HtmlEncode(
            [System.IO.Path]::GetFileName(
                $originalHtmlPath
            )
        )

    $fallbackHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">

<title>ITI365 dashboard generation error</title>

<style>
body{
    margin:0;
    background:#f4f7fb;
    color:#0f1e33;
    font:14px/1.5 "Segoe UI",sans-serif
}

.wrap{
    max-width:900px;
    margin:70px auto;
    padding:0 24px
}

.card{
    background:#fff;
    border:1px solid #e2eaf3;
    border-radius:12px;
    padding:24px;
    box-shadow:0 2px 6px rgb(30 60 120 / .06)
}

h1{
    margin-top:0;
    font-size:22px
}

.error{
    padding:14px;
    border-radius:8px;
    background:#fde2e2;
    color:#991b1b;
    white-space:pre-wrap
}

.btn{
    display:inline-block;
    margin-top:16px;
    padding:9px 13px;
    border-radius:8px;
    background:#2563eb;
    color:#fff;
    text-decoration:none;
    font-weight:700
}

.path{
    margin-top:16px;
    color:#5e7292;
    font-family:Consolas,monospace;
    overflow-wrap:anywhere
}
</style>

</head>

<body>

<div class="wrap">
    <div class="card">

        <h1>
            Dashboard post-processing did not complete
        </h1>

        <p>
            The ITI365 assessment itself completed and the original report is available.
        </p>

        <div class="error">
            $safeMessage
        </div>

        <a
            class="btn"
            href="$safeRawName"
        >
            Open original report
        </a>

        <div class="path">
            Diagnostic log: Dashboard-Generation.log
        </div>

    </div>
</div>

</body>
</html>
"@

    $fallbackHtml |
        Set-Content `
            -Path $dashboardPath `
            -Encoding utf8

    Write-Error "The original ITI365 report was generated, but the modern dashboard post-processing failed. Diagnostic log: $generationLogPath. $($failure.Exception.Message)"
}

$passedCount = @(
    $normalizedTests |
        Where-Object Status -eq 'Passed'
).Count

$failedCount = @(
    $normalizedTests |
        Where-Object Status -eq 'Failed'
).Count

$investigateCount = @(
    $normalizedTests |
        Where-Object Status -eq 'Investigate'
).Count

$skippedCount = @(
    $normalizedTests |
        Where-Object {
            $_.Status -in @(
                'Skipped',
                'NotRun'
            )
        }
).Count

$errorCount = @(
    $normalizedTests |
        Where-Object Status -eq 'Error'
).Count

Write-Host ''

Write-Host `
    'Dashboard successfully generated.' `
    -ForegroundColor Green

Write-Host "  Total       : $($normalizedTests.Count)"

Write-Host `
    "  Passed      : $passedCount" `
    -ForegroundColor Green

Write-Host `
    "  Failed      : $failedCount" `
    -ForegroundColor Red

Write-Host `
    "  Investigate : $investigateCount" `
    -ForegroundColor Magenta

Write-Host `
    "  Skipped     : $skippedCount" `
    -ForegroundColor Yellow

Write-Host `
    "  Errors      : $errorCount" `
    -ForegroundColor Red

Write-Host ''

Write-Host `
    "Modern dashboard : $dashboardPath" `
    -ForegroundColor Cyan

Write-Host `
    "Original report  : $originalHtmlPath"

Write-Host `
    "Structured JSON  : $jsonPath"

if ($OpenReport) {
    Invoke-Item `
        -Path $dashboardPath
}

[pscustomobject]@{
    RunFolder       = $runFolder
    DashboardHtml   = $dashboardPath
    OriginalHtml    = $originalHtmlPath
    JsonReport      = $jsonPath
    Total           = $normalizedTests.Count
    Passed          = $passedCount
    Failed          = $failedCount
    Investigate     = $investigateCount
    SkippedOrNotRun = $skippedCount
    Errors          = $errorCount
}
