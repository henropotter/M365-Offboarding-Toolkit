$ErrorActionPreference = 'Stop'

$publicPath  = Join-Path -Path $PSScriptRoot -ChildPath 'Public'
$privatePath = Join-Path -Path $PSScriptRoot -ChildPath 'Private'

$public  = @(Get-ChildItem -Path (Join-Path -Path $publicPath  -ChildPath '*.ps1') -ErrorAction SilentlyContinue)
$private = @(Get-ChildItem -Path (Join-Path -Path $privatePath -ChildPath '*.ps1') -ErrorAction SilentlyContinue)

foreach ($file in @($public + $private)) {
    try {
        . $file.FullName
    }
    catch {
        Write-Error "Failed to import function file $($file.FullName): $_"
    }
}

Export-ModuleMember -Function $public.BaseName
