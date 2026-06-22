function Connect-OffboardingExchange {
    <#
    .SYNOPSIS
        Connects to Exchange Online once, on first proven need, and records the connection on the session context.

    .DESCRIPTION
        Establishes an Exchange Online connection using the authentication mode of the supplied session context.
        Interactive sessions connect interactively. Application sessions connect app-only using the AppId,
        CertificateThumbprint, and Organization that Connect-OffboardingSession stashed on the context.

        The function is idempotent. If the context already shows Exchange connected, or a live Exchange connection
        already exists, it returns without reconnecting. It sets ExchangeConnected to true on the context once a
        connection is in place.

        The ExchangeOnlineManagement module is installed automatically for the current user if it is missing.
        Only if it cannot be installed, for example on a host with no Gallery access, does the function decline to
        connect and return false, leaving ExchangeConnected false so the caller can record that the mailbox and
        mail-enabled group steps cannot run and escalate accordingly.

    .PARAMETER Context
        The session context returned by Connect-OffboardingSession. Updated in place.

    .EXAMPLE
        if (Connect-OffboardingExchange -Context $session) { ... mailbox steps ... }

    .OUTPUTS
        Boolean. True if Exchange Online is connected, false if it could not be connected.

    .NOTES
        Compatible with Windows PowerShell 5.1 and PowerShell 7. Requires the ExchangeOnlineManagement module for
        the connection to succeed.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Context
    )

    if ($Context.ExchangeConnected) {
        Write-Verbose 'Exchange Online is already connected for this session.'
        return $true
    }

    if (-not $Context.ExchangeModuleAvailable) {
        try {
            $null = Install-OffboardingModule -Name 'ExchangeOnlineManagement'
            $Context.ExchangeModuleAvailable = $true
        }
        catch {
            Write-Warning "ExchangeOnlineManagement is not installed and could not be installed, so mailbox and mail-enabled group steps cannot run: $($_.Exception.Message)"
            return $false
        }
    }

    Import-Module -Name 'ExchangeOnlineManagement' -ErrorAction Stop

    $existing = $null
    try { $existing = Get-ConnectionInformation -ErrorAction SilentlyContinue } catch { $existing = $null }
    if ($existing) {
        Write-Verbose 'Reusing the existing Exchange Online connection.'
        $Context.ExchangeConnected = $true
        return $true
    }

    try {
        if ($Context.AuthMode -eq 'Application') {
            $exo = $Context.ExchangeAuth
            if ([string]::IsNullOrWhiteSpace([string]$exo.Organization)) {
                throw 'Exchange Online application authentication requires the Organization value, which was not supplied to Connect-OffboardingSession.'
            }
            Write-Verbose 'Connecting to Exchange Online using certificate application authentication.'
            Connect-ExchangeOnline -AppId $exo.AppId -CertificateThumbprint $exo.CertificateThumbprint -Organization $exo.Organization -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Write-Verbose 'Connecting to Exchange Online using interactive authentication.'
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Exchange Online connection failed: $($_.Exception.Message)"
        return $false
    }

    $Context.ExchangeConnected = $true
    return $true
}
