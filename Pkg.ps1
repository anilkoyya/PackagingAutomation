[CmdletBinding()]
param(
    [string]$SourceDir = (Split-Path $MyInvocation.MyCommand.Path -Parent),
    [string]$OutputDir = "$env:ProgramData\PackageForge\VCRedist",
    [string]$LogDir    = "$env:ProgramData\PackageForge\Logs\VCRedist-2015-2022"
)
$ErrorActionPreference = 'Stop'

# --- Force 64-bit PowerShell on 64-bit OS (avoids WoW64 reg redirection) -----
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Write-Warning "Relaunching in 64-bit PowerShell..."
    $native = Join-Path $env:WinDir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    & $native -ExecutionPolicy Bypass -File $PSCommandPath @PSBoundParameters
    exit $LASTEXITCODE
}

foreach ($d in $OutputDir,$LogDir) {
    if (-not (Test-Path $d)) { New-Item $d -ItemType Directory -Force | Out-Null }
}

# Catches: Redistributable, Minimum Runtime, Additional Runtime,
# Debug Runtime — every MSI the bootstrapper can drop.
$NameRegex = 'Microsoft Visual C\+\+ 20\d{2}(-20\d{2})?\s+(Redistributable|Minimum Runtime|Additional Runtime|Debug Runtime)'

function Get-VCRedistEntries {
    $out = New-Object System.Collections.Generic.List[object]
    $views = @(
        @{ View=[Microsoft.Win32.RegistryView]::Registry64; Arch='x64' },
        @{ View=[Microsoft.Win32.RegistryView]::Registry32; Arch='x86' }
    )
    foreach ($v in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine, $v.View)
        $un = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
        if (-not $un) { continue }
        foreach ($n in $un.GetSubKeyNames()) {
            $k = $un.OpenSubKey($n); if (-not $k) { continue }
            $dn = $k.GetValue('DisplayName')
            if ($dn -match $NameRegex) {
                $out.Add([PSCustomObject]@{
                    ProductCode    = $n         # subkey name == MSI ProductCode
                    DisplayName    = $dn
                    DisplayVersion = $k.GetValue('DisplayVersion')
                    Publisher      = $k.GetValue('Publisher')
                    InstallDate    = $k.GetValue('InstallDate')
                    InstallSource  = $k.GetValue('InstallSource')
                    SystemComponent= [int]($k.GetValue('SystemComponent', 0))
                    Architecture   = $v.Arch
                })
            }
            $k.Close()
        }
        $un.Close(); $base.Close()
    }
    return $out
}

function Invoke-Bootstrapper {
    param([string]$Exe, [string]$Arch)
    $log = Join-Path $LogDir "VC_${Arch}_install.log"
    Write-Host "[$Arch] Installing $(Split-Path $Exe -Leaf) ..." -ForegroundColor Cyan
    $p = Start-Process -FilePath $Exe `
        -ArgumentList '/install','/quiet','/norestart',"/log `"$log`"" `
        -Wait -PassThru
    Write-Host "[$Arch] Exit code: $($p.ExitCode)"
    if ($p.ExitCode -notin 0,3010,1638) {
        throw "[$Arch] failed ($($p.ExitCode)). See $log"
    }
    return $p.ExitCode
}

# 1. Snapshot
Write-Host "`n[1/4] Pre-install snapshot..." -ForegroundColor Green
$before = @(Get-VCRedistEntries)
Write-Host "      Existing VC++ entries: $($before.Count)"

# 2. Install
Write-Host "`n[2/4] Installing bootstrappers..." -ForegroundColor Green
$x64Exe = Join-Path $SourceDir 'VC_redist.x64.exe'
$x86Exe = Join-Path $SourceDir 'VC_redist.x86.exe'
if (-not (Test-Path $x64Exe)) { throw "Missing: $x64Exe" }
if (-not (Test-Path $x86Exe)) { throw "Missing: $x86Exe" }

$rc64 = Invoke-Bootstrapper -Exe $x64Exe -Arch 'x64'
$rc86 = Invoke-Bootstrapper -Exe $x86Exe -Arch 'x86'
Start-Sleep -Seconds 5

# 3. Re-snapshot
Write-Host "`n[3/4] Post-install snapshot..." -ForegroundColor Green
$after = @(Get-VCRedistEntries)

# 4. Report
Write-Host "`n[4/4] All VC++ 2015-2022 MSIs on this system:" -ForegroundColor Green
$after | Sort-Object Architecture, DisplayName |
    Format-Table ProductCode, Architecture, DisplayName, DisplayVersion -AutoSize -Wrap

$new = $after | Where-Object { $_.ProductCode -notin ($before.ProductCode) }
if ($new) {
    Write-Host "`nInstalled by THIS run:" -ForegroundColor Cyan
    $new | Format-Table ProductCode, Architecture, DisplayName, DisplayVersion -AutoSize -Wrap
} else {
    Write-Host "`n(no new entries — same/newer version was already present)" -ForegroundColor DarkYellow
}

# Sanity check
$cnt64 = ($after | Where-Object Architecture -eq 'x64').Count
$cnt86 = ($after | Where-Object Architecture -eq 'x86').Count
Write-Host "`nFound: x64=$cnt64  x86=$cnt86" -ForegroundColor Magenta
if ($cnt64 -eq 0) {
    Write-Warning "No x64 entries detected — check that VC_redist.x64.exe actually ran (see $LogDir)."
}

# Persist
$jsonPath = Join-Path $OutputDir 'product_codes.json'
$csvPath  = Join-Path $OutputDir 'product_codes.csv'
[ordered]@{
    Package            = 'VCRedist-2015-2022'
    CapturedOn         = (Get-Date).ToString('o')
    Host               = $env:COMPUTERNAME
    PSBitness          = if ([Environment]::Is64BitProcess) { '64-bit' } else { '32-bit' }
    InstallerExitCodes = @{ x64 = $rc64; x86 = $rc86 }
    AllProducts        = $after
    NewProducts        = $new
} | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding UTF8
$after | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nWritten:`n  $jsonPath`n  $csvPath" -ForegroundColor Green
