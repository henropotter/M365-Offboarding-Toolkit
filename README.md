# M365 Offboarding Toolkit

Environment-adaptive Microsoft 365 user offboarding for hybrid and cloud-only estates.

> Status: the read-only assessment is built and usable now through `Invoke-OffboardingAssessment`. The execute phase, and the safety and approval layer that gates it, are in active development. See Status and roadmap below.

## What it does

One approved user in, a consistent offboard out, with the judgment calls escalated rather than left to whoever ran the tool. The toolkit detects the environment it runs in and adapts:

- Hybrid: on-prem Active Directory synced to Entra ID, with Exchange Online for mail.
- Cloud-only: Entra ID with Exchange Online, no on-prem directory.

It runs in two phases. ASSESS reads everything, decides every action by policy, prints the plan, and changes nothing. EXECUTE performs the actions, but only after explicit approval. The assessment phase is what ships today.

## Requirements

- PowerShell 5.1 or 7
- Microsoft.Graph.Authentication

Optional, enabling the matching paths when present:

- ExchangeOnlineManagement for the mailbox and mail-enabled group steps
- ActiveDirectory and ADSync for the on-prem disable, move, synced-group removal, and delta sync steps

The PowerShell Gallery modules above install themselves on first use. ActiveDirectory and ADSync are detected only and are present on a domain controller or an Entra Connect server, so the on-prem detail is most accurate when the toolkit is run there.

## Install

Clone the repository and import the module:

```powershell
git clone https://github.com/henropotter/M365-Offboarding-Toolkit.git
cd M365-Offboarding-Toolkit
Import-Module ./M365OffboardingToolkit.psd1
```

## Usage

Assess one person. This signs in interactively, requests read-only permissions, reads the account, groups, licenses, and mailbox, and prints the plan. It changes nothing:

```powershell
Invoke-OffboardingAssessment -UserPrincipalName 'jordan.avery@contoso.com'
```

Because the interactive sign-in asks for read-only permissions, the consent screen lists Read access only and the session has no ability to make changes.

Sample output:

```
Offboarding plan for Jordan Avery <jordan.avery@contoso.com>

Account
  Disable the account
     - Currently enabled: True
     - Account is cloud-only, so it is disabled directly in Entra ID.

Sign-in sessions
  Revoke all active sign-in sessions
     - Signs the user out of Microsoft 365 on every device and forces a new sign-in, which then fails because the account is disabled.

Groups
  Remove group memberships (5 total)
     - Sales Team  (cloud group, removed through Graph)
     - All Company DL  (distribution list, removed through Exchange)
     - Project Falcon  (Microsoft 365 group, removed through Exchange)
     - West Region  (dynamic group, skipped because membership clears on its own when the account is disabled)
     - Box Users  (grants access to data, held until owned content is dealt with first)

Licenses
  Reclaim 2 license(s)
     - Reclaim license: Microsoft 365 E5
     - Reclaim license: Power BI Pro
     - Assigned through a group, drops automatically on group exit: Visio Plan 2

Mailbox
  Convert to shared and remove delegated access
     - Mailbox size: about 7.17 GB
     - Convert from a user mailbox to a shared mailbox, so the contents stay reachable without a license.
     - Remove Full Access for: alex.kim@contoso.com
     - Remove Send As for: alex.kim@contoso.com
     - Remove calendar delegate access for: taylor.reed@contoso.com
```

Return the plan as an object instead of printing it, for further processing:

```powershell
$plan = Invoke-OffboardingAssessment -UserPrincipalName 'jordan.avery@contoso.com' -PassThru
```

Run unattended with a certificate application registration instead of an interactive sign-in:

```powershell
Invoke-OffboardingAssessment -UserPrincipalName 'jordan.avery@contoso.com' `
    -ClientId $clientId -TenantId $tenantId `
    -CertificateThumbprint $thumbprint -Organization 'contoso.onmicrosoft.com'
```

## How a person is assessed

- Account: whether it is cloud-only or synced from on-prem, which decides where it gets disabled.
- Groups: each membership is classified by how it would be handled. Cloud groups through Graph, distribution lists and Microsoft 365 groups through Exchange, on-prem groups in Active Directory, and dynamic groups left alone because they clear on their own.
- Licenses: directly assigned licenses are separated from group-assigned ones, since the latter drop automatically when the user leaves the group.
- Mailbox: type, size, and the Full Access, Send As, Send on Behalf, and calendar delegates, read from Exchange Online when a mailbox is present.
- Holds: a litigation or retention hold flips the plan to preserve the mailbox and licenses rather than reclaim them.

## Status and roadmap

Built and usable now:

- Read-only assessment: sign-in, environment detection, group and license and mailbox reads, and the printed plan.

In active development:

- Handling for data-app and direct application entitlements, so access like Box is held rather than removed.
- A safety and accountability layer: legal-hold and privileged-user gates, and a typed approval step.
- The execute phase that carries out an approved plan.
- Continuous integration with PSScriptAnalyzer.

## License

MIT. See [LICENSE](LICENSE).
