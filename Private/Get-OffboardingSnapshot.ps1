function Get-OffboardingSnapshot {
    <#
    .SYNOPSIS
        Reads one user from Microsoft Graph (and Exchange Online when needed) and returns a snapshot for planning. Read-only.

    .DESCRIPTION
        Builds the user snapshot that Get-OffboardingPlan consumes, entirely from read operations. It resolves the
        user, classifies their group memberships by how each would be handled, and separates directly assigned
        licenses from group-assigned ones. When the Exchange-requirement test signals a mailbox, it connects
        Exchange Online read-only and reads the mailbox type and size, the Full Access, Send As, Send on Behalf,
        and calendar delegates, and whether the mailbox is on a litigation or retention hold.

        Nothing here changes the directory or the mailbox. It only reads. The result is a plain object that the
        plan builder turns into the human-readable plan.

        Group classification in this version is structural (dynamic, on-prem synced, Microsoft 365, distribution
        list, or cloud security). Detection of data-app holds and direct enterprise-app entitlements is a later
        layer and is not performed here.

    .PARAMETER Context
        The session context returned by Connect-OffboardingSession. Used to connect Exchange on demand.

    .PARAMETER UserPrincipalName
        The user principal name of the person to assess.

    .EXAMPLE
        $snapshot = Get-OffboardingSnapshot -Context $session -UserPrincipalName 'jordan.avery@contoso.com'

    .OUTPUTS
        PSCustomObject shaped for Get-OffboardingPlan, plus a MailboxAssessed flag.

    .NOTES
        Compatible with Windows PowerShell 5.1 and PowerShell 7. Performs only read operations.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Context,

        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    $select = 'id,displayName,userPrincipalName,accountEnabled,onPremisesSyncEnabled,mail,proxyAddresses,assignedPlans,assignedLicenses,licenseAssignmentStates'
    $userUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=$select"
    $u = Invoke-MgGraphRequest -Method GET -Uri $userUri -ErrorAction Stop

    $groups = @()
    $groupUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/memberOf`?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,onPremisesSyncEnabled&`$top=999"
    while ($groupUri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $groupUri -ErrorAction Stop
        if ($resp.value) {
            foreach ($item in $resp.value) {
                if ($item.'@odata.type' -eq '#microsoft.graph.group') { $groups += $item }
            }
        }
        $groupUri = $resp.'@odata.nextLink'
    }

    $groupInfo = @()
    foreach ($g in $groups) {
        $types = @($g.groupTypes)
        if ($types -contains 'DynamicMembership') { $route = 'Dynamic' }
        elseif ($g.onPremisesSyncEnabled) { $route = 'AD' }
        elseif ($types -contains 'Unified') { $route = 'M365' }
        elseif ($g.mailEnabled) { $route = 'DL' }
        else { $route = 'Cloud' }
        $groupInfo += [pscustomobject]@{ Name = [string]$g.displayName; Route = $route }
    }

    $skuMap = @{}
    try {
        $skus = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuId,skuPartNumber' -ErrorAction Stop
        if ($skus.value) { foreach ($s in $skus.value) { $skuMap[[string]$s.skuId] = [string]$s.skuPartNumber } }
    }
    catch { }

    $states = @($u.licenseAssignmentStates)
    $directIds = @($states | Where-Object { -not $_.assignedByGroup } | ForEach-Object { [string]$_.skuId } | Where-Object { $_ } | Select-Object -Unique)
    $groupIds = @($states | Where-Object { $_.assignedByGroup } | ForEach-Object { [string]$_.skuId } | Where-Object { $_ } | Select-Object -Unique)

    $directLicenses = @($directIds | ForEach-Object { if ($skuMap.ContainsKey($_)) { $skuMap[$_] } else { $_ } })
    $groupLicenses = @($groupIds | ForEach-Object { if ($skuMap.ContainsKey($_)) { $skuMap[$_] } else { $_ } })

    $mailbox = $null
    $onHold = $false
    $mailboxAssessed = $false

    $need = Test-OffboardingExchangeRequirement -User $u -Group $groups
    if ($need.Required) {
        $exo = $false
        try { $exo = Connect-OffboardingExchange -Context $Context } catch { $exo = $false }
        if ($exo) {
            try {
                $mbx = Get-EXOMailbox -Identity $UserPrincipalName -Properties RecipientTypeDetails, LitigationHoldEnabled, InPlaceHolds, ComplianceTagHoldApplied, GrantSendOnBehalfTo -ErrorAction Stop

                $sizeMB = 0
                try {
                    $stat = Get-EXOMailboxStatistics -Identity $UserPrincipalName -ErrorAction Stop
                    $raw = "$($stat.TotalItemSize)"
                    if ($raw -match '\(([\d,]+) bytes\)') {
                        $bytes = [int64](($matches[1] -replace ',', ''))
                        $sizeMB = [math]::Round($bytes / 1MB, 0)
                    }
                }
                catch { }

                $fullAccess = @()
                try { $fullAccess = @(Get-EXOMailboxPermission -Identity $UserPrincipalName -ErrorAction Stop | Where-Object { "$($_.User)" -notlike 'NT AUTHORITY\SELF' -and -not $_.IsInherited -and ("$($_.AccessRights)" -match 'FullAccess') } | ForEach-Object { "$($_.User)" }) } catch { }

                $sendAs = @()
                try { $sendAs = @(Get-EXORecipientPermission -Identity $UserPrincipalName -ErrorAction Stop | Where-Object { "$($_.Trustee)" -notlike 'NT AUTHORITY\SELF' -and ("$($_.AccessRights)" -match 'SendAs') } | ForEach-Object { "$($_.Trustee)" }) } catch { }

                $sendOnBehalf = @()
                if ($mbx.GrantSendOnBehalfTo) { $sendOnBehalf = @($mbx.GrantSendOnBehalfTo | ForEach-Object { "$_" }) }

                $calendar = @()
                try {
                    $calFolder = $UserPrincipalName + ':\Calendar'
                    $calendar = @(Get-EXOMailboxFolderPermission -Identity $calFolder -ErrorAction Stop | Where-Object { "$($_.User)" -notmatch '^(Default|Anonymous)$' -and ("$($_.AccessRights)" -notmatch 'None') } | ForEach-Object { "$($_.User)" })
                }
                catch { }

                $type = 'UserMailbox'
                if ($mbx.RecipientTypeDetails -eq 'SharedMailbox') { $type = 'SharedMailbox' }

                $onHold = ([bool]$mbx.LitigationHoldEnabled) -or (@($mbx.InPlaceHolds).Count -gt 0) -or ([bool]$mbx.ComplianceTagHoldApplied)

                $mailbox = [pscustomobject]@{
                    Exists            = $true
                    Type              = $type
                    SizeMB            = $sizeMB
                    FullAccess        = $fullAccess
                    SendAs            = $sendAs
                    SendOnBehalf      = $sendOnBehalf
                    CalendarDelegates = $calendar
                }
                $mailboxAssessed = $true
            }
            catch {
                $mailbox = [pscustomobject]@{ Exists = $false; Type = 'None'; SizeMB = 0; FullAccess = @(); SendAs = @(); SendOnBehalf = @(); CalendarDelegates = @() }
                $mailboxAssessed = $true
            }
        }
    }
    else {
        $mailbox = [pscustomobject]@{ Exists = $false; Type = 'None'; SizeMB = 0; FullAccess = @(); SendAs = @(); SendOnBehalf = @(); CalendarDelegates = @() }
        $mailboxAssessed = $true
    }

    [pscustomobject]@{
        DisplayName       = [string]$u.displayName
        UserPrincipalName = [string]$u.userPrincipalName
        AccountEnabled    = [bool]$u.accountEnabled
        OnPremSynced      = [bool]$u.onPremisesSyncEnabled
        OnHold            = $onHold
        Groups            = $groupInfo
        DirectLicenses    = $directLicenses
        GroupLicenses     = $groupLicenses
        Mailbox           = $mailbox
        MailboxAssessed   = $mailboxAssessed
    }
}
