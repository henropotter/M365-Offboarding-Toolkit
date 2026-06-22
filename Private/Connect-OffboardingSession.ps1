function Connect-OffboardingSession {
    <#
    .SYNOPSIS
        Establishes the Microsoft Graph session for an offboarding run and detects which surfaces this host can reach.

    .DESCRIPTION
        Connects to Microsoft Graph in one of two modes and returns a single context object that every later
        offboarding step keys its behavior off.

        Interactive (default): an operator signs in as themselves. Graph uses delegated permissions, so the run
        can only do what that operator's own roles allow. Intended for hands-on, single-user runs.

        Application (certificate): supply ClientId, TenantId, and CertificateThumbprint to connect app-only with
        no interactive prompt. The work then runs under the application's identity using admin-consented
        application permissions. Intended for unattended execution or a back-end service.

        The function also detects whether the ActiveDirectory and ADSync modules are present, which determines
        whether the on-prem hybrid path is possible on this host, and whether the ExchangeOnlineManagement module
        is present, which determines whether the mailbox and mail-enabled-group steps can run later. Exchange is
        not connected here. It is connected on first proven need by Connect-OffboardingExchange, using the
        connection details this function stashes on the returned context.

        The required Microsoft.Graph.Authentication module is installed automatically for the current user if it
        is missing. ActiveDirectory and ADSync are never installed: they are not Gallery modules, and their
        absence simply means this host cannot perform the on-prem steps.

    .PARAMETER ClientId
        Application (client) ID of the app registration to authenticate as. Supplying this, with TenantId and
        CertificateThumbprint, selects certificate application-only authentication.

    .PARAMETER TenantId
        Directory (tenant) ID to connect to. Required for application authentication.

    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate in the local certificate store used to authenticate the app registration.
        Required for application authentication.

    .PARAMETER Organization
        The tenant's primary domain, for example contoso.onmicrosoft.com. Used only when Exchange Online is later
        connected in application mode, where it is required. Ignored in interactive mode.

    .PARAMETER Scopes
        Delegated Graph permission scopes to request in interactive mode. Defaults to the set the offboarding tool
        needs end to end. Ignored in application mode, where permissions come from the app registration's consent.

    .EXAMPLE
        $session = Connect-OffboardingSession

        Connects interactively in the operator's own context and returns the session context.

    .EXAMPLE
        $session = Connect-OffboardingSession -ClientId $appId -TenantId $tenant -CertificateThumbprint $thumb -Organization 'contoso.onmicrosoft.com'

        Connects app-only by certificate, suitable for unattended runs.

    .OUTPUTS
        PSCustomObject describing the session: AuthMode, TenantId, Account, GraphConnected, ExchangeConnected,
        ExchangeModuleAvailable, ActiveDirectoryAvailable, AdSyncAvailable, the stashed ExchangeAuth details, and
        CreatedAt.

    .NOTES
        Compatible with Windows PowerShell 5.1 and PowerShell 7. Requires the Microsoft.Graph.Authentication module.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(ParameterSetName = 'Application', Mandatory = $true)]
        [string]$ClientId,

        [Parameter(ParameterSetName = 'Application', Mandatory = $true)]
        [string]$TenantId,

        [Parameter(ParameterSetName = 'Application', Mandatory = $true)]
        [string]$CertificateThumbprint,

        [Parameter(ParameterSetName = 'Application')]
        [string]$Organization,

        [Parameter(ParameterSetName = 'Interactive')]
        [string[]]$Scopes = @(
            'User.ReadWrite.All',
            'Group.ReadWrite.All',
            'GroupMember.ReadWrite.All',
            'Directory.Read.All',
            'Organization.Read.All',
            'Application.Read.All',
            'AppRoleAssignment.ReadWrite.All'
        )
    )

    $isAppAuth = $PSCmdlet.ParameterSetName -eq 'Application'

    $null = Install-OffboardingModule -Name 'Microsoft.Graph.Authentication'
    Import-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction Stop

    $graphContext = $null
    try { $graphContext = Get-MgContext } catch { $graphContext = $null }

    if ($null -eq $graphContext) {
        if ($isAppAuth) {
            Write-Verbose 'Connecting to Microsoft Graph using certificate application authentication.'
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        }
        else {
            Write-Verbose 'Connecting to Microsoft Graph using interactive authentication.'
            Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
        }
        try { $graphContext = Get-MgContext } catch { $graphContext = $null }
    }
    else {
        Write-Verbose 'Reusing the existing Microsoft Graph connection.'
    }

    if ($null -eq $graphContext) {
        throw 'Microsoft Graph connection could not be established.'
    }

    $authMode = 'Interactive'
    if ($graphContext.AuthType -eq 'AppOnly') { $authMode = 'Application' }

    $account = $graphContext.Account
    if ([string]::IsNullOrWhiteSpace([string]$account)) { $account = $graphContext.AppName }

    $exoAppId = $null
    $exoThumbprint = $null
    if ($isAppAuth) {
        $exoAppId = $ClientId
        $exoThumbprint = $CertificateThumbprint
    }

    $exchangeAuth = [pscustomobject]@{
        Mode                  = $authMode
        AppId                 = $exoAppId
        CertificateThumbprint = $exoThumbprint
        Organization          = $Organization
    }

    [pscustomobject]@{
        AuthMode                 = $authMode
        TenantId                 = $graphContext.TenantId
        Account                  = $account
        GraphConnected           = $true
        ExchangeConnected        = $false
        ExchangeModuleAvailable  = [bool](Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')
        ExchangeAuth             = $exchangeAuth
        ActiveDirectoryAvailable = [bool](Get-Module -ListAvailable -Name 'ActiveDirectory')
        AdSyncAvailable          = [bool](Get-Module -ListAvailable -Name 'ADSync')
        CreatedAt                = (Get-Date)
    }
}
