function Get-OffboardingPlan {
    <#
    .SYNOPSIS
        Builds a structured offboarding plan for one person from a snapshot of their account, changing nothing.

    .DESCRIPTION
        Takes a snapshot describing a single user (their account, group memberships, licenses, and mailbox) and
        returns a plan object: an ordered list of the actions an offboard would take, each with the specific
        details underneath. This is the read-only half of the tool. It decides what would happen and never does it.

        The snapshot is whatever the fetch layer assembles from Microsoft Graph and Exchange Online. Keeping the
        plan builder separate from the fetching means it can be run and tested with sample data, with no tenant.

    .PARAMETER Snapshot
        A user snapshot. Expected shape:
          DisplayName, UserPrincipalName, AccountEnabled, OnPremSynced, OnHold,
          Groups          (each item has Name and Route, where Route is one of
                           Cloud, AD, DL, M365, Dynamic, Hold),
          DirectLicenses  (friendly names), GroupLicenses (friendly names),
          Mailbox         (Exists, Type, SizeMB, FullAccess, SendAs, SendOnBehalf, CalendarDelegates)

    .EXAMPLE
        $plan = Get-OffboardingPlan -Snapshot $snapshot

    .OUTPUTS
        PSCustomObject with Person (string), Steps (array of step objects), and Notes (string array).

    .NOTES
        Pure function with no external calls, so it can be tested without a live tenant.
        Compatible with Windows PowerShell 5.1 and PowerShell 7.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        $Snapshot
    )

    $steps = [System.Collections.Generic.List[object]]::new()

    # Account
    $enabledLine = "Currently enabled: $($Snapshot.AccountEnabled)"
    if ($Snapshot.OnPremSynced) {
        $whereLine = 'Account is synced from on-prem Active Directory, so it is disabled in AD and the change syncs up to Entra ID.'
    }
    else {
        $whereLine = 'Account is cloud-only, so it is disabled directly in Entra ID.'
    }
    $steps.Add([pscustomobject]@{
            Category = 'Account'
            Action   = 'Disable the account'
            Details  = @($enabledLine, $whereLine)
            Flag     = 'Change'
        })

    # Sign-in sessions
    $steps.Add([pscustomobject]@{
            Category = 'Sign-in sessions'
            Action   = 'Revoke all active sign-in sessions'
            Details  = @('Signs the user out of Microsoft 365 on every device and forces a new sign-in, which then fails because the account is disabled.')
            Flag     = 'Change'
        })

    # Groups
    $groups = @($Snapshot.Groups)
    if ($groups.Count -eq 0) {
        $steps.Add([pscustomobject]@{
                Category = 'Groups'
                Action   = 'No group memberships to remove'
                Details  = @()
                Flag     = 'Skip'
            })
    }
    else {
        $groupLines = foreach ($g in $groups) {
            switch ("$($g.Route)") {
                'AD'      { "$($g.Name)  (on-prem group, removed in Active Directory then synced)" }
                'DL'      { "$($g.Name)  (distribution list, removed through Exchange)" }
                'M365'    { "$($g.Name)  (Microsoft 365 group, removed through Exchange)" }
                'Cloud'   { "$($g.Name)  (cloud group, removed through Graph)" }
                'Dynamic' { "$($g.Name)  (dynamic group, skipped because membership clears on its own when the account is disabled)" }
                'Hold'    { "$($g.Name)  (grants access to data, held until owned content is dealt with first)" }
                default   { "$($g.Name)  (removed)" }
            }
        }
        $steps.Add([pscustomobject]@{
                Category = 'Groups'
                Action   = "Remove group memberships ($($groups.Count) total)"
                Details  = @($groupLines)
                Flag     = 'Change'
            })
    }

    # Licenses
    if ($Snapshot.OnHold) {
        $steps.Add([pscustomobject]@{
                Category = 'Licenses'
                Action   = 'Keep all licenses'
                Details  = @('A legal or retention hold is in place, so licenses are kept to preserve the mailbox. Nothing is removed.')
                Flag     = 'Hold'
            })
    }
    else {
        $direct = @($Snapshot.DirectLicenses)
        $viaGroup = @($Snapshot.GroupLicenses)
        $licenseLines = @()
        if ($direct.Count -gt 0) {
            foreach ($l in $direct) { $licenseLines += "Reclaim license: $l" }
        }
        else {
            $licenseLines += 'No directly assigned licenses to reclaim.'
        }
        foreach ($l in $viaGroup) { $licenseLines += "Assigned through a group, drops automatically on group exit: $l" }

        if ($direct.Count -gt 0) { $licenseAction = "Reclaim $($direct.Count) license(s)" } else { $licenseAction = 'No licenses to reclaim' }
        $steps.Add([pscustomobject]@{
                Category = 'Licenses'
                Action   = $licenseAction
                Details  = @($licenseLines)
                Flag     = 'Change'
            })
    }

    # Mailbox
    $mbx = $Snapshot.Mailbox
    if (-not $mbx -or -not $mbx.Exists) {
        $steps.Add([pscustomobject]@{
                Category = 'Mailbox'
                Action   = 'No mailbox to handle'
                Details  = @('This user has no Exchange Online mailbox.')
                Flag     = 'Skip'
            })
    }
    else {
        $mbxLines = @()
        $sizeGb = [math]::Round((([double]$mbx.SizeMB) / 1024), 2)
        $mbxLines += "Mailbox size: about $sizeGb GB"
        if ("$($mbx.Type)" -eq 'SharedMailbox') {
            $mbxLines += 'Already a shared mailbox, so no conversion is needed.'
        }
        else {
            $mbxLines += 'Convert from a user mailbox to a shared mailbox, so the contents stay reachable without a license.'
        }
        $fullAccess = @($mbx.FullAccess)
        $sendAs = @($mbx.SendAs)
        $sendOnBehalf = @($mbx.SendOnBehalf)
        $calendar = @($mbx.CalendarDelegates)
        if ($fullAccess.Count -gt 0) { $mbxLines += "Remove Full Access for: $($fullAccess -join ', ')" } else { $mbxLines += 'No Full Access delegates to remove.' }
        if ($sendAs.Count -gt 0) { $mbxLines += "Remove Send As for: $($sendAs -join ', ')" } else { $mbxLines += 'No Send As delegates to remove.' }
        if ($sendOnBehalf.Count -gt 0) { $mbxLines += "Remove Send on Behalf for: $($sendOnBehalf -join ', ')" }
        if ($calendar.Count -gt 0) { $mbxLines += "Remove calendar delegate access for: $($calendar -join ', ')" } else { $mbxLines += 'No calendar delegates to remove.' }
        if ($Snapshot.OnHold) { $mbxLines += 'Mailbox is on hold, so it stays licensed and is never deleted.' }

        $steps.Add([pscustomobject]@{
                Category = 'Mailbox'
                Action   = 'Convert to shared and remove delegated access'
                Details  = @($mbxLines)
                Flag     = 'Change'
            })
    }

    # Notes
    $notes = @()
    if ($Snapshot.OnHold) {
        $notes += 'This person is on a legal or retention hold. Licenses and mailbox are preserved and nothing is deleted. Confirm records disposition with the owner before closing.'
    }

    [pscustomobject]@{
        Person = "$($Snapshot.DisplayName) <$($Snapshot.UserPrincipalName)>"
        Steps  = $steps.ToArray()
        Notes  = @($notes)
    }
}
