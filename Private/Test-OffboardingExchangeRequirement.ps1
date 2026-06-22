function Test-OffboardingExchangeRequirement {
    <#
    .SYNOPSIS
        Decides whether Exchange Online must be connected to offboard a given user, and explains why.

    .DESCRIPTION
        Reads signals already available from Microsoft Graph and returns whether the run needs Exchange Online.
        Exchange is required when any of the following is true:

        1. The user has an Exchange Online service plan whose status is not Deleted. The non-Deleted states
           (Enabled, Warning, Suspended, LockedOut) all mean a mailbox exists or its data is being preserved,
           which is what catches a recently de-licensed mailbox during its retention window. A Deleted plan is
           not used as a trigger on its own, because Graph keeps historical service-plan entries that may no
           longer reflect the current license.

        2. The user still carries mailbox attributes, namely a populated mail value or an SMTP proxy address.
           This is an independent signal that survives the license-removal window.

        3. The user belongs to a mail-enabled group that Graph cannot edit, namely a classic distribution list
           or a mail-enabled security group. Microsoft 365 (Unified) group membership is editable through Graph
           and does not require Exchange.

        The license plan is only a trigger to connect. Once Exchange is connected, Exchange itself is the source
        of truth for whether a mailbox exists and what is done with it.

    .PARAMETER User
        The user object as returned by Microsoft Graph. Expected to include assignedPlans, mail, and proxyAddresses.

    .PARAMETER Group
        The user's group memberships as returned by Microsoft Graph. Each item is expected to include mailEnabled
        and groupTypes. May be empty.

    .EXAMPLE
        Test-OffboardingExchangeRequirement -User $user -Group $groups

    .OUTPUTS
        PSCustomObject with Required (boolean) and Reasons (string array).

    .NOTES
        Pure decision function with no external calls, so it can be unit tested without a live tenant.
        Compatible with Windows PowerShell 5.1 and PowerShell 7.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        $User,

        [Parameter()]
        [object[]]$Group = @()
    )

    $reasons = [System.Collections.Generic.List[string]]::new()

    foreach ($plan in @($User.assignedPlans)) {
        $service = "$($plan.service)".ToLower()
        $status = "$($plan.capabilityStatus)"
        if ($service -eq 'exchange' -and $status -ne 'Deleted') {
            $reasons.Add("Exchange Online service plan present (status: $status).")
            break
        }
    }

    $hasMail = -not [string]::IsNullOrWhiteSpace([string]$User.mail)
    $hasSmtpProxy = $false
    foreach ($proxy in @($User.proxyAddresses)) {
        if ("$proxy".ToLower().StartsWith('smtp:')) {
            $hasSmtpProxy = $true
            break
        }
    }
    if ($hasMail -or $hasSmtpProxy) {
        $reasons.Add('Residual mailbox attributes present (mail or SMTP proxy address).')
    }

    foreach ($g in @($Group)) {
        $mailEnabled = [bool]$g.mailEnabled
        $isUnified = $false
        foreach ($type in @($g.groupTypes)) {
            if ("$type" -eq 'Unified') {
                $isUnified = $true
                break
            }
        }
        if ($mailEnabled -and -not $isUnified) {
            $name = $g.displayName
            if ([string]::IsNullOrWhiteSpace([string]$name)) { $name = $g.id }
            $reasons.Add("Member of a mail-enabled group Graph cannot edit: $name.")
        }
    }

    [pscustomobject]@{
        Required = ($reasons.Count -gt 0)
        Reasons  = $reasons.ToArray()
    }
}
