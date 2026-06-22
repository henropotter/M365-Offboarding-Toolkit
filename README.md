# M365 Offboarding Toolkit

Environment-adaptive Microsoft 365 user offboarding for hybrid and cloud-only estates.

> Status: work in progress. This is the module scaffold. Offboarding logic lands in the next iterations.

## What it does

One approved user in, a consistent offboard out, with the judgment calls escalated rather than left to whoever ran the tool. The toolkit detects the environment it runs in and adapts:

- Hybrid: on-prem Active Directory synced to Entra ID, with Exchange Online for mail.
- Cloud-only: Entra ID with Exchange Online, no on-prem directory.

It runs in two phases. ASSESS reads everything, decides every action by policy, prints the plan, and changes nothing. EXECUTE performs the actions, but only after explicit approval.

A full description of the safety model and the accountability outputs will be documented here as the tool is built.

## Requirements

- PowerShell 5.1 or 7
- Microsoft.Graph.Authentication

Optional, enabling the matching paths when present:

- ExchangeOnlineManagement for the mailbox and mail-enabled group steps
- ActiveDirectory and ADSync for the on-prem disable, move, synced-group removal, and delta sync steps

## Install

To be documented.

## Usage

To be documented.

## License

MIT. See [LICENSE](LICENSE).
