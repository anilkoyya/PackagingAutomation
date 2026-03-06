# =============================================================================
# Package-AppV-Local.ps1
# App-V 5.0 Sequencing — Local execution (no remote VM)
# Runs directly on the host machine or inside a Docker/Sandbox container
# =============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AppName,
    [string]$AppVersion = "1.0",
    [string]$InstallCommand,
    [string]$InstallArgs,
    [string]$ExclusionPaths,
    [string]$OutputName,
    [Parameter(Mandatory)][string]$WorkDir,
    [Parameter(Mandatory)][string]$MediaDir,
    [Parameter(Mandatory)][string]$OutputDir,
    [string]$SequencerPath = "C:\Program Files\Microsoft Application Virtualization\Sequencer\AppvSequencer.exe"
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $OutputDir "appv-packaging.log"

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$ts] $msg" | Tee-Object -FilePath $logFile -Append
}

Write-Log "══════════════════════════════════════════════"
Write-Log "  App-V 5.0 Sequencing (Local Mode)"
Write-Log "  App: $AppName v$AppVersion"
Write-Log "══════════════════════════════════════════════"

if (-not $OutputName) {
    $OutputName = "${AppName}_${AppVersion}" -replace '[^\w\-\.]', '_'
}

# Ensure directories exist
New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null
New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null

# ─── Step 1: Locate Installer ────────────────────────────────────────────────
Write-Log "[1/5] Locating installer..."

$installerPath = ""
if ($InstallCommand) {
    $fullPath = Join-Path $MediaDir $InstallCommand
    if (Test-Path $fullPath) {
        $installerPath = $fullPath
    } else {
        $found = Get-ChildItem $MediaDir -Recurse -File |
                 Where-Object { $_.Name -like "*$InstallCommand*" } |
                 Select-Object -First 1
        if ($found) { $installerPath = $found.FullName }
    }
}
if (-not $installerPath) {
    $found = Get-ChildItem $MediaDir -Recurse -Include "*.msi","*.exe" | Select-Object -First 1
    if ($found) { $installerPath = $found.FullName }
    else { throw "No installer found in $MediaDir" }
}

Write-Log "  Installer: $installerPath"

# ─── Step 2: Generate Sequencer Template ─────────────────────────────────────
Write-Log "[2/5] Generating sequencer template..."

$exclusions = @(
    "C:\Windows\Temp",
    "C:\Users\*\AppData\Local\Temp",
    "C:\Users\*\AppData\Local\Microsoft\Windows\INetCache"
)
if ($ExclusionPaths) {
    $extras = $ExclusionPaths -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $exclusions += $extras
}

$templatePath = Join-Path $WorkDir "SequencerTemplate.appvt"
$templateXml = @"
<?xml version="1.0" encoding="utf-8"?>
<SequencerTemplate xmlns="http://schemas.microsoft.com/appv/2013/sequencertemplate">
  <Package>
    <Name>$AppName</Name>
    <Version>$AppVersion</Version>
    <DisplayName>$AppName $AppVersion</DisplayName>
    <Description>Auto-sequenced: $AppName $AppVersion</Description>
  </Package>
  <Installer>
    <PrimaryVirtualApplicationDirectory>C:\Program Files\$AppName</PrimaryVirtualApplicationDirectory>
  </Installer>
  <ExcludedItems>
$($exclusions | ForEach-Object { "    <ExcludedItem>$_</ExcludedItem>" } | Out-String)
  </ExcludedItems>
  <Options>
    <GenerateMSI>true</GenerateMSI>
    <AllowCOMObjects>true</AllowCOMObjects>
    <EnforceSecurityDescriptors>true</EnforceSecurityDescriptors>
  </Options>
</SequencerTemplate>
"@

$templateXml | Out-File -FilePath $templatePath -Encoding utf8
Write-Log "  Template: $templatePath"

# ─── Step 3: Run Sequencing ──────────────────────────────────────────────────
Write-Log "[3/5] Running App-V Sequencer..."

$isMSI = $installerPath -match '\.msi$'

# Method A: Use App-V PowerShell module (preferred)
if (Get-Command "New-AppvSequencerPackage" -ErrorAction SilentlyContinue) {
    Write-Log "  Using App-V PowerShell module..."

    $seqArgs = @{
        Name = $OutputName
        TemplateFilePath = $templatePath
        OutputPath = $OutputDir
        PrimaryVirtualApplicationDirectory = "C:\Program Files\$AppName"
        Installer = $installerPath
    }
    if ($InstallArgs) {
        $seqArgs["InstallerArguments"] = $InstallArgs
    }

    New-AppvSequencerPackage @seqArgs

# Method B: Use sequencer CLI
} elseif (Test-Path $SequencerPath) {
    Write-Log "  Using Sequencer CLI..."

    $seqArgs = @(
        "/Template:`"$templatePath`"",
        "/Installer:`"$installerPath`"",
        "/OutputPath:`"$OutputDir`"",
        "/PackageName:`"$OutputName`"",
        "/InstallPath:`"C:\Program Files\$AppName`""
    )
    if ($InstallArgs) {
        $seqArgs += "/InstallerArguments:`"$InstallArgs`""
    }

    $proc = Start-Process -FilePath $SequencerPath `
        -ArgumentList $seqArgs `
        -Wait -PassThru -NoNewWindow

    if ($proc.ExitCode -ne 0) {
        Write-Log "  WARNING: Sequencer exited with code $($proc.ExitCode)"
    }

# Method C: Manual capture (snapshot-diff approach)
} else {
    Write-Log "  No App-V Sequencer found — using manual snapshot-diff approach..."
    Write-Log "  Taking pre-install filesystem snapshot..."

    $snapshotFile = Join-Path $WorkDir "pre-snapshot.json"
    $watchDirs = @("C:\Program Files", "C:\Program Files (x86)", "C:\ProgramData")

    $preSnapshot = @{}
    foreach ($dir in $watchDirs) {
        if (Test-Path $dir) {
            $files = Get-ChildItem $dir -Recurse -ErrorAction SilentlyContinue |
                     Select-Object FullName, Length, LastWriteTime
            $preSnapshot[$dir] = $files
        }
    }

    # Take registry snapshot
    $regPre = @{}
    @("HKLM:\SOFTWARE", "HKCU:\SOFTWARE") | ForEach-Object {
        $items = Get-ChildItem $_ -Recurse -ErrorAction SilentlyContinue |
                 Select-Object PSPath
        $regPre[$_] = $items
    }

    Write-Log "  Installing application: $installerPath $InstallArgs"

    if ($isMSI) {
        $installArgs2 = if ($InstallArgs) { $InstallArgs } else { "/qn /norestart ALLUSERS=1" }
        $proc = Start-Process -FilePath "msiexec.exe" `
            -ArgumentList "/i `"$installerPath`" $installArgs2" `
            -Wait -PassThru
    } else {
        $proc = Start-Process -FilePath $installerPath `
            -ArgumentList $InstallArgs `
            -Wait -PassThru
    }

    Write-Log "  Install exit code: $($proc.ExitCode)"
    Start-Sleep -Seconds 5

    # Post-install snapshot diff
    Write-Log "  Computing filesystem diff..."

    $changedFiles = @()
    foreach ($dir in $watchDirs) {
        if (Test-Path $dir) {
            $postFiles = Get-ChildItem $dir -Recurse -ErrorAction SilentlyContinue
            $preFiles = $preSnapshot[$dir]
            if ($preFiles) {
                $prePaths = @($preFiles | ForEach-Object { $_.FullName })
                $newFiles = $postFiles | Where-Object { $_.FullName -notin $prePaths }
                $changedFiles += $newFiles
            } else {
                $changedFiles += $postFiles
            }
        }
    }

    Write-Log "  Found $($changedFiles.Count) new/changed files"

    # Build output report (without actual App-V package since sequencer not available)
    $diffReport = @{
        Application = "$AppName $AppVersion"
        NewFiles = @($changedFiles | Select-Object FullName, Length | ForEach-Object {
            @{ Path = $_.FullName; Size = $_.Length }
        })
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Note = "App-V Sequencer not found. Generated filesystem diff report instead."
    }
    $diffReport | ConvertTo-Json -Depth 4 | Out-File (Join-Path $OutputDir "${OutputName}_diff-report.json") -Encoding utf8
    Write-Log "  Diff report saved"
}

# ─── Step 4: Validate Output ─────────────────────────────────────────────────
Write-Log "[4/5] Validating output..."

$appvFile = Get-ChildItem $OutputDir -Filter "*.appv" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($appvFile) {
    Write-Log "  Package: $($appvFile.Name) ($([math]::Round($appvFile.Length/1MB, 2)) MB)"
} else {
    $jsonReport = Get-ChildItem $OutputDir -Filter "*diff-report.json" | Select-Object -First 1
    if ($jsonReport) {
        Write-Log "  No .appv file (sequencer not available). Diff report: $($jsonReport.Name)"
    } else {
        Write-Log "  WARNING: No output files found"
    }
}

# ─── Step 5: Generate Report ─────────────────────────────────────────────────
Write-Log "[5/5] Generating packaging report..."

$report = @{
    PackageType        = "App-V 5.0"
    ExecutionMode      = "Local"
    ApplicationName    = $AppName
    ApplicationVersion = $AppVersion
    PackageName        = $OutputName
    PackageFile        = if ($appvFile) { $appvFile.Name } else { "N/A (sequencer not found)" }
    PackageSize        = if ($appvFile) { "$([math]::Round($appvFile.Length/1MB, 2)) MB" } else { "N/A" }
    InstallerUsed      = $installerPath
    Timestamp          = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Status             = if ($appvFile) { "Success" } else { "Partial — diff report generated" }
}
$report | ConvertTo-Json -Depth 3 | Out-File (Join-Path $OutputDir "packaging-report.json") -Encoding utf8

Write-Log "══════════════════════════════════════════════"
Write-Log "  App-V Packaging Complete"
Write-Log "══════════════════════════════════════════════"

exit 0
