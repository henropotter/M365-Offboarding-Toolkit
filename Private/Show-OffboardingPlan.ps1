function Show-OffboardingPlan {
    <#
    .SYNOPSIS
        Renders an offboarding plan as readable, indented text.

    .DESCRIPTION
        Takes a plan from Get-OffboardingPlan and returns it as lines of text, one section per planned area, with
        the specific details indented underneath. It returns the lines rather than printing them directly, so the
        same output can be shown on screen, written to a log, or checked in a test.

    .PARAMETER Plan
        A plan object produced by Get-OffboardingPlan.

    .EXAMPLE
        Get-OffboardingPlan -Snapshot $snapshot | Show-OffboardingPlan

    .OUTPUTS
        String, emitted one line at a time.

    .NOTES
        Compatible with Windows PowerShell 5.1 and PowerShell 7.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Plan
    )
    process {
        Write-Output ''
        Write-Output "Offboarding plan for $($Plan.Person)"
        Write-Output ''

        foreach ($step in @($Plan.Steps)) {
            $marker = ''
            if ($step.Flag -eq 'Hold') { $marker = '   [HOLD]' }
            elseif ($step.Flag -eq 'Skip') { $marker = '   [nothing to do]' }

            Write-Output ("{0}{1}" -f $step.Category, $marker)
            Write-Output ("  {0}" -f $step.Action)
            foreach ($detail in @($step.Details)) {
                Write-Output ("     - {0}" -f $detail)
            }
            Write-Output ''
        }

        if (@($Plan.Notes).Count -gt 0) {
            Write-Output 'Notes'
            foreach ($note in @($Plan.Notes)) {
                Write-Output ("     - {0}" -f $note)
            }
            Write-Output ''
        }
    }
}
