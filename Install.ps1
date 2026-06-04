<#
.SYNOPSIS
    Installs VC++ 2015-2022 Redistributable (x64 + x86) and reports the
    ProductCodes of every MSI the bootstrappers registered.

.DESCRIPTION
    1. Snapshots existing VC++ 2015-2022 entries from the Uninstall registry
       (both native and WOW6432Node hives).
    2. Installs VC_redist.x64.exe then VC_redist.x86.exe silently.
    3. Re-snapshots and diffs to surface what THIS run installed.
    4. Writes JSON + CSV for PackageForge metadata ingestion.

    Run elevated. Exit codes handled: 0, 3010 (reboot), 1638 (newer present).
#>
[CmdletBinding()]
param(
    [string]$SourceDir = (Split-Path $MyInvocation.MyCommand.Path -Parent),
    [string]$OutputDir = "$env:ProgramData\PackageForge\VCRedist",
    [string]$LogDir    = "$env:ProgramData\PackageForge\Logs\VCRedist-2015-2022"
)

$ErrorActionPreference = 'Stop'
foreach ($d in $OutputDir,$LogDir) {
    if (-not (Test-Path $d)) { New-Item $d -ItemType Directory -Force | Out-Null }
}

# Both hives so we catch x64 MSIs (native) and x86 MSIs (WOW6432Node view)
$UninstallHives = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

# Matches 2015-2019, 2015-2022, and future 20xx-20xx naming
$NameRegex = 'Microsoft Visual C\+\+ 20\d{2}-20\d{2} Redistributable'

function Get-VCRedistEntries {
    foreach ($hive in $UninstallHives) {
        if (-not (Test-Path $hive)) { continue }
        Get-ChildItem $hive -ErrorAction SilentlyContinue | ForEach-Object {
            $p = $_ | Get-ItemProperty -ErrorAction SilentlyContinue
            if ($p.DisplayName -match $NameRegex) {
                [PSCustomObject]@{
                    ProductCode    = $_.PSChildName     # reg key name == MSI ProductCode
                    DisplayName    = $p.DisplayName
                    DisplayVersion = $p.DisplayVersion
                    Publisher      = $p.Publisher
                    InstallSource  = $p.InstallSource
                    HiveView       = if ($hive -match 'WOW6432Node') { 'x86' } else { 'x64' }
                }
            }
        }
    }
}

function Invoke-Bootstrapper {
    param([string]$Exe, [string]$Arch)
    $log = Join-Path $LogDir "VC_${Arch}_install.log"
    Write-Host "[$Arch] Installing $(Split-Path $Exe -Leaf) ..." -ForegroundColor Cyan
    $proc = Start-Process -FilePath $Exe `
        -ArgumentList '/install','/quiet','/norestart',"/log `"$log`"" `
        -Wait -PassThru
    Write-Host "[$Arch] Exit code: $($proc.ExitCode)" -ForegroundColor Yellow
    if ($proc.ExitCode -notin 0, 3010, 1638) {
        throw "[$Arch] Installer failed ($($proc.ExitCode)). See $log"
    }
    return $proc.ExitCode
}

# -- 1. Pre snapshot ----------------------------------------------------------
Write-Host "`n[1/4] Pre-install snapshot..." -ForegroundColor Green
$before = @(Get-VCRedistEntries)
Write-Host "      Existing entries: $($before.Count)"

# -- 2. Install --------------------------------------------------------------
Write-Host "`n[2/4] Installing bootstrappers..." -ForegroundColor Green
$x64Exe = Join-Path $SourceDir 'VC_redist.x64.exe'
$x86Exe = Join-Path $SourceDir 'VC_redist.x86.exe'
if (-not (Test-Path $x64Exe)) { throw "Missing: $x64Exe" }
if (-not (Test-Path $x86Exe)) { throw "Missing: $x86Exe" }

$rc64 = Invoke-Bootstrapper -Exe $x64Exe -Arch 'x64'
$rc86 = Invoke-Bootstrapper -Exe $x86Exe -Arch 'x86'

Start-Sleep -Seconds 3   # let MSI session unwind

# -- 3. Post snapshot --------------------------------------------------------
Write-Host "`n[3/4] Post-install snapshot..." -ForegroundColor Green
$after = @(Get-VCRedistEntries)

# -- 4. Report ---------------------------------------------------------------
Write-Host "`n[4/4] All registered VC++ 2015-2022 MSIs:" -ForegroundColor Green
$after | Sort-Object DisplayName |
    Format-Table ProductCode, DisplayName, DisplayVersion, HiveView -AutoSize

$newOnes = $after | Where-Object { $_.ProductCode -notin ($before.ProductCode) }
if ($newOnes) {
    Write-Host "`nInstalled by THIS run:" -ForegroundColor Cyan
    $newOnes | Format-Table ProductCode, DisplayName, DisplayVersion -AutoSize
} else {
    Write-Host "`n(no new entries — same/newer version was already present)" -ForegroundColor DarkYellow
}

# -- Persist for PackageForge ------------------------------------------------
$jsonPath = Join-Path $OutputDir 'product_codes.json'
$csvPath  = Join-Path $OutputDir 'product_codes.csv'

[ordered]@{
    Package            = 'VCRedist-2015-2022'
    CapturedOn         = (Get-Date).ToString('o')
    Host               = $env:COMPUTERNAME
    InstallerExitCodes = @{ x64 = $rc64; x86 = $rc86 }
    AllProducts        = $after
    NewProducts        = $newOnes
} | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding UTF8

$after | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nWritten:" -ForegroundColor Green
Write-Host "  $jsonPath"
Write-Host "  $csvPath"
