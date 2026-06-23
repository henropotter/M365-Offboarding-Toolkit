[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName
)

Import-Module (Join-Path $PSScriptRoot 'M365OffboardingToolkit.psd1') -Force
$module = Get-Module M365OffboardingToolkit

$readScopes = @(
    'User.Read.All',
    'Group.Read.All',
    'Directory.Read.All',
    'Organization.Read.All',
    'Application.Read.All'
)

& $module {
    param($upn, $scopes)

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }

    Write-Host 'Connecting read-only. A browser sign-in opens, and the consent screen lists Read permissions only.' -ForegroundColor Cyan
    $session = Connect-OffboardingSession -Scopes $scopes

    Write-Host ''
    Write-Host ('Signed in as {0}   mode {1}   tenant {2}' -f $session.Account, $session.AuthMode, $session.TenantId) -ForegroundColor Green

    $snapshot = Get-OffboardingSnapshot -Context $session -UserPrincipalName $upn

    if (-not $snapshot.MailboxAssessed) {
        Write-Host ''
        Write-Host 'Exchange was not connected, so the mailbox was not assessed. Ignore the mailbox section below.' -ForegroundColor Yellow
    }

    Get-OffboardingPlan -Snapshot $snapshot | Show-OffboardingPlan
} $UserPrincipalName $readScopes
