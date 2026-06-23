function Invoke-OffboardingAssessment {
    <#
    .SYNOPSIS
        Runs a read-only offboarding assessment for one user and shows the plan. Changes nothing.

    .DESCRIPTION
        The public entry point for assessing a person. It connects to Microsoft 365 read-only, reads the account,
        groups, licenses, and mailbox, builds the offboarding plan, and prints it. It never disables, removes, or
        changes anything. This is the safe, look-only half of the toolkit.

        By default it signs in interactively in a browser and requests read-only permissions, so the consent screen
        lists Read access only and the session has no ability to make changes. For an unattended run, supply the
        application sign-in details (ClientId, TenantId, CertificateThumbprint, and Organization) and it connects as
        an app registration instead. Either way, the assessment performs only read operations.

    .PARAMETER UserPrincipalName
        The user principal name of the person to assess, for example jordan.avery@contoso.com. Accepts pipeline input.

    .PARAMETER ClientId
        App registration (client) ID, for certificate app-only sign-in. Supply with TenantId and CertificateThumbprint.

    .PARAMETER TenantId
        Directory (tenant) ID, for certificate app-only sign-in.

    .PARAMETER CertificateThumbprint
        Certificate thumbprint in the local store, for certificate app-only sign-in.

    .PARAMETER Organization
        Tenant primary domain, for example contoso.onmicrosoft.com, used for app-only Exchange sign-in.

    .PARAMETER PassThru
        Return the structured plan object instead of printing the readable plan.

    .EXAMPLE
        Invoke-OffboardingAssessment -UserPrincipalName 'jordan.avery@contoso.com'

        Signs in interactively read-only and prints the plan for one person.

    .EXAMPLE
        Invoke-OffboardingAssessment -UserPrincipalName 'jordan.avery@contoso.com' -ClientId $id -TenantId $tid -CertificateThumbprint $thumb -Organization 'contoso.onmicrosoft.com'

        Runs unattended using a certificate app registration.

    .EXAMPLE
        $plan = Invoke-OffboardingAssessment -UserPrincipalName 'jordan.avery@contoso.com' -PassThru

        Returns the plan object for further processing instead of printing it.

    .OUTPUTS
        By default, the rendered plan as text. With -PassThru, the structured plan object.

    .NOTES
        Performs only read operations. Compatible with Windows PowerShell 5.1 and PowerShell 7.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$UserPrincipalName,

        [Parameter()]
        [string]$ClientId,

        [Parameter()]
        [string]$TenantId,

        [Parameter()]
        [string]$CertificateThumbprint,

        [Parameter()]
        [string]$Organization,

        [Parameter()]
        [switch]$PassThru
    )
    process {
        $readScopes = @(
            'User.Read.All',
            'Group.Read.All',
            'Directory.Read.All',
            'Organization.Read.All',
            'Application.Read.All'
        )

        if ($ClientId -and $TenantId -and $CertificateThumbprint) {
            Write-Verbose 'Connecting with certificate application sign-in.'
            $session = Connect-OffboardingSession -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -Organization $Organization
        }
        else {
            Write-Verbose 'Connecting interactively, read-only.'
            $session = Connect-OffboardingSession -Scopes $readScopes
        }

        Write-Verbose ("Signed in as {0} (mode {1}, tenant {2})." -f $session.Account, $session.AuthMode, $session.TenantId)

        $snapshot = Get-OffboardingSnapshot -Context $session -UserPrincipalName $UserPrincipalName

        if (-not $snapshot.MailboxAssessed) {
            Write-Warning 'Exchange was not connected, so the mailbox was not assessed. The mailbox section reflects that.'
        }

        $plan = Get-OffboardingPlan -Snapshot $snapshot

        if ($PassThru) {
            $plan
        }
        else {
            $plan | Show-OffboardingPlan
        }
    }
}
