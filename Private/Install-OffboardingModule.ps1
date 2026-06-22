function Install-OffboardingModule {
    <#
    .SYNOPSIS
        Ensures a PowerShell Gallery module is available, installing it for the current user if it is missing.

    .DESCRIPTION
        Checks whether the named module is already present, optionally at or above a minimum version, and returns
        immediately if it is. If it is missing, the module is installed from the PowerShell Gallery into the
        current user's scope, with no elevation required.

        This is used only for modules the Gallery actually publishes, namely Microsoft.Graph.Authentication and
        ExchangeOnlineManagement. The ActiveDirectory and ADSync modules are deliberately never passed here. They
        are not Gallery modules: ActiveDirectory ships with the Windows RSAT feature and ADSync is installed by
        Microsoft Entra Connect, and both exist only on a domain-joined or Entra Connect host. Their absence is a
        host-capability signal that the on-prem path is not available here, not a defect to fix by installing.

        If installation fails, for example on a host with no Gallery access, the function throws a clear,
        actionable error rather than letting a later step fail with an obscure one.

    .PARAMETER Name
        The Gallery module name to ensure is available.

    .PARAMETER MinimumVersion
        Optional minimum version. If the installed module is older, a newer version is installed.

    .EXAMPLE
        Install-OffboardingModule -Name 'Microsoft.Graph.Authentication'

    .OUTPUTS
        Boolean. True when the module is available after the call.

    .NOTES
        Compatible with Windows PowerShell 5.1 and PowerShell 7.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter()]
        [string]$MinimumVersion
    )

    $present = @(Get-Module -ListAvailable -Name $Name)
    if ($MinimumVersion) {
        $present = @($present | Where-Object { $_.Version -ge [version]$MinimumVersion })
    }
    if ($present.Count -gt 0) {
        Write-Verbose "Module $Name is already available."
        return $true
    }

    Write-Verbose "Module $Name is not installed. Installing from the PowerShell Gallery for the current user."

    # Windows PowerShell 5.1 may negotiate a TLS version the Gallery rejects, so raise it to TLS 1.2.
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    catch {
        Write-Verbose 'Could not adjust the TLS protocol; continuing with the current setting.'
    }

    # A fresh Windows PowerShell host may lack the NuGet provider, which would otherwise prompt interactively.
    try {
        $null = Get-PackageProvider -Name 'NuGet' -ForceBootstrap -ErrorAction Stop
    }
    catch {
        Write-Verbose 'NuGet provider bootstrap was not required or not possible; continuing.'
    }

    $installParams = @{
        Name         = $Name
        Scope        = 'CurrentUser'
        Force        = $true
        AllowClobber = $true
        ErrorAction  = 'Stop'
    }
    if ($MinimumVersion) {
        $installParams['MinimumVersion'] = $MinimumVersion
    }

    try {
        Install-Module @installParams
    }
    catch {
        throw "Required module $Name could not be installed automatically: $($_.Exception.Message) Install it manually with: Install-Module $Name -Scope CurrentUser"
    }

    return ([bool]@(Get-Module -ListAvailable -Name $Name).Count)
}
