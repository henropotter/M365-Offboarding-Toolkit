@{
    RootModule           = 'M365OffboardingToolkit.psm1'
    ModuleVersion        = '0.1.0'
    GUID                 = '353c4a51-1789-4124-8a1d-aea8e757baf1'
    Author               = 'henropotter'
    CompanyName          = 'Community'
    Copyright            = '(c) 2026 henropotter. Released under the MIT License.'
    Description          = 'Environment-adaptive Microsoft 365 user offboarding for hybrid (on-prem Active Directory plus Entra ID) and cloud-only (Entra ID plus Exchange Online) estates. Detects the environment, then runs a two-phase assess-and-execute offboard with safety checks and an accountability record.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredModules      = @('Microsoft.Graph.Authentication')
    FunctionsToExport    = @('Invoke-OffboardingAssessment')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData = @{
        PSData = @{
            Tags       = @('Entra', 'EntraID', 'Microsoft365', 'M365', 'Offboarding', 'IAM', 'ExchangeOnline', 'ActiveDirectory', 'Hybrid', 'IdentityLifecycle')
            LicenseUri = 'https://github.com/henropotter/M365-Offboarding-Toolkit/blob/main/LICENSE'
            ProjectUri = 'https://github.com/henropotter/M365-Offboarding-Toolkit'
        }
    }
}
