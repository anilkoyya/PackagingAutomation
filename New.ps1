<#
.SYNOPSIS
    Extracts MSIs from VC_redist bootstrappers and dumps ProductCode / 
    ProductVersion / UpgradeCode for inclusion in the install wrapper.
.NOTES
    Bootstrapper extracts to /layout <dir> revealing:
      packages\vcRuntimeMinimum_<arch>\vc_runtimeMinimum_<arch>.msi
      packages\vcRuntimeAdditional_<arch>\vc_runtimeAdditional_<arch>.msi
#>
[CmdletBinding()]
param(
    [string]$SourceDir = (Split-Path $MyInvocation.MyCommand.Path -Parent),
    [string]$WorkDir   = "$env:TEMP\VCRedist_Extract"
)

$ErrorActionPreference = 'Stop'
if (Test-Path $WorkDir) { Remove-Item $WorkDir -Recurse -Force }
New-Item $WorkDir -ItemType Directory | Out-Null

function Get-MsiProperty {
    param([string]$MsiPath, [string[]]$Properties)
    $wi  = New-Object -ComObject WindowsInstaller.Installer
    $db  = $wi.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$wi,@($MsiPath,0))
    $out = [ordered]@{ File = Split-Path $MsiPath -Leaf }
    foreach ($p in $Properties) {
        $q  = "SELECT Value FROM Property WHERE Property='$p'"
        $v  = $db.GetType().InvokeMember('OpenView','InvokeMethod',$null,$db,($q))
        $v.GetType().InvokeMember('Execute','InvokeMethod',$null,$v,$null) | Out-Null
        $r  = $v.GetType().InvokeMember('Fetch','InvokeMethod',$null,$v,$null)
        $out[$p] = if ($r) { $r.GetType().InvokeMember('StringData','GetProperty',$null,$r,1) } else { $null }
    }
    [PSCustomObject]$out
}

foreach ($arch in 'x64','x86') {
    $exe = Join-Path $SourceDir "VC_redist.$arch.exe"
    if (-not (Test-Path $exe)) { Write-Warning "Missing: $exe"; continue }

    $layout = Join-Path $WorkDir $arch
    Write-Host "Extracting $exe ..." -ForegroundColor Cyan
    Start-Process $exe -ArgumentList "/layout `"$layout`" /quiet" -Wait

    Get-ChildItem $layout -Recurse -Filter *.msi | ForEach-Object {
        Get-MsiProperty -MsiPath $_.FullName `
            -Properties 'ProductCode','ProductVersion','ProductName','UpgradeCode'
    } | Format-Table -AutoSize
}
