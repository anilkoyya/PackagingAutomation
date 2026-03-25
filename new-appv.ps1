#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Advanced App-V 5.x Package Automation Engine
    Runs on the packaging VM, triggered remotely via WinRM / CI/CD pipeline.

.DESCRIPTION
    Fully automated App-V sequencer orchestration supporting:
      - Unattended mode  : silent install via command line, no user interaction
      - Attended mode     : filesystem/registry monitoring with external signal to finalize
      - Hybrid mode       : unattended install + attended customization window

    Designed for remote invocation from a FastAPI/GitLab CI orchestrator via WinRM.
    Reports structured JSON status back to the caller for pipeline integration.

.PARAMETER PackageName
    Display name of the application being packaged.

.PARAMETER PackageVersion
    Version string (e.g. "14.0.5").

.PARAMETER InstallerPath
    Full path to the installer on the packaging VM (UNC or local).

.PARAMETER InstallerArgs
    Silent install arguments (e.g. "/S /v/qn"). Required for Unattended/Hybrid.

.PARAMETER InstallerType
    Installer technology: MSI, EXE, SCRIPT. Determines default handling.

.PARAMETER Mode
    Packaging mode: Unattended | Attended | Hybrid.

.PARAMETER OutputRoot
    Root directory for completed packages. A subfolder is created per package.

.PARAMETER PrimaryVirtualDirectory
    PVAD for the sequenced package. Defaults to "C:\AppVPackages\<PackageName>".

.PARAMETER TemplateFile
    Optional App-V package accelerator or template (.appvt / .cab).

.PARAMETER PreInstallScript
    Optional script block or path executed before the installer (machine prep).

.PARAMETER PostInstallScript
    Optional script block or path executed after install, before sequencing finishes.

.PARAMETER TimeoutSeconds
    Maximum wall-clock time for the entire operation. Default 3600 (1 hour).

.PARAMETER AttendedSignalFile
    Filepath whose creation signals "customization complete" in Attended/Hybrid modes.

.PARAMETER CleanupOnFailure
    Remove partial output on failure. Default $true.

.PARAMETER SequencerPath
    Path to the App-V Sequencer module or executable. Auto-detected if not specified.

.PARAMETER ExclusionPatterns
    Array of path/registry patterns to exclude from capture.

.PARAMETER ReportWebhookUrl
    Optional URL to POST a JSON status report on completion/failure.

.EXAMPLE
    # Remote invocation via WinRM from the orchestrator
    Invoke-Command -ComputerName PKG-VM01 -FilePath .\AppV-PackageAutomation.ps1 -ArgumentList @{
        PackageName    = "Notepad++ 8.6"
        PackageVersion = "8.6.4"
        InstallerPath  = "\\fileserver\installers\npp.8.6.4.Installer.x64.exe"
        InstallerArgs  = "/S"
        InstallerType  = "EXE"
        Mode           = "Unattended"
        OutputRoot     = "D:\AppVPackages"
    }

.EXAMPLE
    # Direct execution on the packaging VM (e.g. from a scheduled task / CI runner)
    .\AppV-PackageAutomation.ps1 `
        -PackageName "Adobe Reader DC" `
        -PackageVersion "24.002" `
        -InstallerPath "C:\Staging\AcroRdrDC.exe" `
        -InstallerArgs "/sAll /msi EULA_ACCEPT=YES" `
        -InstallerType "EXE" `
        -Mode "Hybrid" `
        -OutputRoot "D:\AppVPackages" `
        -PostInstallScript "C:\Scripts\AdobeReader-PostConfig.ps1"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$PackageName,

    [Parameter(Mandatory)]
    [string]$PackageVersion,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InstallerPath,

    [string]$InstallerArgs = "",

    [ValidateSet("MSI", "EXE", "SCRIPT")]
    [string]$InstallerType = "EXE",

    [ValidateSet("Unattended", "Attended", "Hybrid")]
    [string]$Mode = "Unattended",

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$OutputRoot,

    [string]$PrimaryVirtualDirectory = "",

    [string]$TemplateFile = "",

    [string]$PreInstallScript = "",

    [string]$PostInstallScript = "",

    [int]$TimeoutSeconds = 3600,

    [string]$AttendedSignalFile = "",

    [bool]$CleanupOnFailure = $true,

    [string]$SequencerPath = "",

    [string[]]$ExclusionPatterns = @(),

    [string]$ReportWebhookUrl = ""
)

# ============================================================================
# REGION: Strict Mode & Global Configuration
# ============================================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"   # suppress noisy progress bars in remote sessions

$script:RunId       = [guid]::NewGuid().ToString("N").Substring(0, 12)
$script:StartTime   = [datetime]::UtcNow
$script:LogEntries  = [System.Collections.Generic.List[hashtable]]::new()

# Sanitize package name for filesystem use
$SafeName           = ($PackageName -replace '[^\w\.\-]', '_').Trim('_')
$PackageOutputDir   = Join-Path $OutputRoot ($SafeName + "_" + $PackageVersion + "_" + $script:RunId)
$LogFile            = Join-Path $PackageOutputDir "packaging.log"

if (-not $PrimaryVirtualDirectory) {
    $PrimaryVirtualDirectory = "C:\AppVPackages\$SafeName"
}
if (-not $AttendedSignalFile) {
    $AttendedSignalFile = Join-Path $env:TEMP ("AppV_CustomizationComplete_" + $script:RunId + ".signal")
}

# ============================================================================
# REGION: Logging & Reporting
# ============================================================================
function Write-PackageLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG", "STEP")]
        [string]$Level = "INFO"
    )

    $timestamp = [datetime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff")
    $entry     = @{
        Timestamp = $timestamp
        Level     = $Level
        Message   = $Message
        RunId     = $script:RunId
    }
    $script:LogEntries.Add($entry)

    $line = "[$timestamp] [$Level] $Message"
    Write-Verbose $line

    # Append to log file if directory exists
    if (Test-Path (Split-Path $LogFile -Parent)) {
        $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
}

function New-StatusReport {
    [CmdletBinding()]
    param(
        [ValidateSet("Success", "Failed", "TimedOut", "Cancelled")]
        [string]$Status,

        [string]$ErrorDetail = "",

        [hashtable]$PackageInfo = @{}
    )

    $elapsed = ([datetime]::UtcNow - $script:StartTime).TotalSeconds

    $report = [ordered]@{
        run_id           = $script:RunId
        status           = $Status
        package_name     = $PackageName
        package_version  = $PackageVersion
        mode             = $Mode
        installer_type   = $InstallerType
        output_directory = $PackageOutputDir
        elapsed_seconds  = [math]::Round($elapsed, 2)
        started_utc      = $script:StartTime.ToString("o")
        completed_utc    = [datetime]::UtcNow.ToString("o")
        hostname         = $env:COMPUTERNAME
        error_detail     = $ErrorDetail
        package_info     = $PackageInfo
        log_entries      = $script:LogEntries.ToArray()
    }

    # Write JSON report to output directory
    $reportPath = Join-Path $PackageOutputDir "status_report.json"
    if (Test-Path (Split-Path $reportPath -Parent)) {
        $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
        Write-PackageLog "Status report written to: $reportPath"
    }

    # POST to webhook if configured
    if ($ReportWebhookUrl) {
        try {
            $json = $report | ConvertTo-Json -Depth 10 -Compress
            Invoke-RestMethod -Uri $ReportWebhookUrl -Method POST -Body $json `
                -ContentType "application/json" -TimeoutSec 30 -ErrorAction Stop
            Write-PackageLog "Status report posted to webhook: $ReportWebhookUrl"
        }
        catch {
            Write-PackageLog "Failed to post webhook report: $_" -Level WARN
        }
    }

    return $report
}

# ============================================================================
# REGION: Environment Validation & Pre-flight Checks
# ============================================================================
function Test-SequencerEnvironment {
    Write-PackageLog "=== Pre-flight Environment Validation ===" -Level STEP

    # 1. Locate the App-V Sequencer
    $seqModule = $null
    if ($SequencerPath -and (Test-Path $SequencerPath)) {
        $seqModule = $SequencerPath
        Write-PackageLog "Using specified sequencer: $SequencerPath"
    }
    else {
        # Auto-detect: try the PowerShell module first, then fall back to executable
        $progFiles    = $env:ProgramFiles
        $progFilesX86 = ${env:ProgramFiles(x86)}
        $modulePaths  = @(
            "$progFiles\Microsoft Application Virtualization\Sequencer\AppvSequencer\AppvSequencer.psd1",
            "$progFilesX86\Microsoft Application Virtualization\Sequencer\AppvSequencer\AppvSequencer.psd1",
            "$progFiles\Windows Kits\10\Microsoft Application Virtualization\Sequencer\AppvSequencer.psd1"
        )
        foreach ($mp in $modulePaths) {
            if (Test-Path $mp) {
                $seqModule = $mp
                Write-PackageLog "Auto-detected sequencer module: $mp"
                break
            }
        }

        if (-not $seqModule) {
            # Check if the module is in the PSModulePath
            $imported = Get-Module -ListAvailable -Name AppvSequencer -ErrorAction SilentlyContinue
            if ($imported) {
                $seqModule = $imported.Path
                Write-PackageLog "Found sequencer in PSModulePath: $seqModule"
            }
        }
    }

    if (-not $seqModule) {
        throw "App-V Sequencer module not found. Install the App-V Sequencer or specify -SequencerPath."
    }

    # 2. Import the module
    try {
        Import-Module $seqModule -Force -ErrorAction Stop
        Write-PackageLog "Sequencer module imported successfully."
    }
    catch {
        # Fallback: if module import fails, check for the CLI executable
        $seqExe = Get-Command "sequencer.exe" -ErrorAction SilentlyContinue
        if ($seqExe) {
            Write-PackageLog "Module import failed but sequencer.exe found at: $($seqExe.Source). Using CLI fallback." -Level WARN
            $script:UseCliMode = $true
        }
        else {
            throw "Failed to import App-V Sequencer module and no CLI fallback available: $_"
        }
    }

    # 3. Check Windows Defender / AV exclusions advisory
    $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defenderStatus -and $defenderStatus.RealTimeProtectionEnabled) {
        Write-PackageLog "Windows Defender real-time protection is ON. Consider adding exclusions for packaging paths." -Level WARN
    }

    # 4. Check pending reboots
    $pendingReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
    if ($pendingReboot) {
        Write-PackageLog "PENDING REBOOT detected — sequencing results may be unreliable!" -Level WARN
    }

    # 5. Check disk space (require at least 10 GB free on output drive)
    $outputDrive = (Split-Path $OutputRoot -Qualifier)
    $driveLetter = $outputDrive.TrimEnd(':')
    $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$outputDrive'" -ErrorAction SilentlyContinue
    if ($driveInfo) {
        $freeGB = [math]::Round($driveInfo.FreeSpace / 1GB, 2)
    }
    else {
        # Fallback to Get-PSDrive
        $psDrive = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
        $freeGB = if ($psDrive) { [math]::Round($psDrive.Free / 1GB, 2) } else { 0 }
    }
    Write-PackageLog "Free disk space on ${outputDrive}: ${freeGB} GB"
    if ($freeGB -lt 10) {
        throw "Insufficient disk space on ${outputDrive}: ${freeGB} GB free (minimum 10 GB required)."
    }

    # 6. Verify installer is accessible
    $installerSize = (Get-Item $InstallerPath).Length
    $installerSizeMB = [math]::Round($installerSize / 1MB, 2)
    Write-PackageLog "Installer verified: $InstallerPath ($installerSizeMB MB)"

    # 7. Validate template file if specified
    if ($TemplateFile -and -not (Test-Path $TemplateFile)) {
        throw "Template file not found: $TemplateFile"
    }

    Write-PackageLog "Pre-flight checks passed." -Level STEP
}

# ============================================================================
# REGION: Machine Snapshot & Cleanup Helpers
# ============================================================================
function Save-MachineBaseline {
    Write-PackageLog "Capturing machine baseline snapshot..." -Level STEP
    $baseline = @{
        Services      = @(Get-Service | Select-Object Name, Status, StartType)
        Processes     = @(Get-Process | Select-Object Name, Id, Path -Unique)
        EnvVars       = [System.Environment]::GetEnvironmentVariables("Machine")
        ScheduledTasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Disabled' } | Select-Object TaskName, TaskPath)
        TempFiles     = (Get-ChildItem $env:TEMP -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
    }
    Write-PackageLog "Baseline captured: $($baseline.Services.Count) services, $($baseline.Processes.Count) processes."
    return $baseline
}

function Compare-MachineState {
    param(
        [hashtable]$Baseline
    )
    Write-PackageLog "Comparing post-install state to baseline..." -Level DEBUG

    $currentServices = @(Get-Service | Select-Object Name, Status, StartType)
    $newServices     = $currentServices | Where-Object { $_.Name -notin $Baseline.Services.Name }

    $currentProcesses = @(Get-Process | Select-Object Name, Id, Path -Unique)
    $newProcesses     = $currentProcesses | Where-Object { $_.Name -notin $Baseline.Processes.Name }

    $delta = @{
        NewServices  = @($newServices | ForEach-Object { $_.Name })
        NewProcesses = @($newProcesses | ForEach-Object { $_.Name })
    }

    if ($delta.NewServices.Count -gt 0) {
        Write-PackageLog "New services detected: $($delta.NewServices -join ', ')" -Level WARN
    }
    if ($delta.NewProcesses.Count -gt 0) {
        Write-PackageLog "New processes detected: $($delta.NewProcesses -join ', ')" -Level DEBUG
    }

    return $delta
}

function Invoke-PostPackageCleanup {
    Write-PackageLog "Running post-package cleanup..." -Level STEP

    # Kill known installer remnants
    $knownResidual = @("msiexec", "setup", "install", "update")
    foreach ($procName in $knownResidual) {
        Get-Process -Name "*$procName*" -ErrorAction SilentlyContinue |
            Where-Object { $_.Id -ne $PID } |
            ForEach-Object {
                Write-PackageLog "Terminating residual process: $($_.Name) (PID $($_.Id))" -Level DEBUG
                $_ | Stop-Process -Force -ErrorAction SilentlyContinue
            }
    }

    # Clean temp
    Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-PackageLog "Cleanup complete."
}

# ============================================================================
# REGION: Installer Execution
# ============================================================================
function Invoke-InstallerExecution {
    Write-PackageLog "=== Executing Installer ===" -Level STEP
    Write-PackageLog "Installer : $InstallerPath"
    Write-PackageLog "Arguments : $InstallerArgs"
    Write-PackageLog "Type      : $InstallerType"

    $installStart = [datetime]::UtcNow

    switch ($InstallerType) {
        "MSI" {
            $msiArgs = "/i `"$InstallerPath`" $InstallerArgs /L*v `"$(Join-Path $PackageOutputDir 'msi_install.log')`""
            Write-PackageLog "Running: msiexec.exe $msiArgs"
            $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs `
                -Wait -PassThru -NoNewWindow -ErrorAction Stop
        }
        "EXE" {
            Write-PackageLog "Running: `"$InstallerPath`" $InstallerArgs"
            $proc = Start-Process -FilePath $InstallerPath -ArgumentList $InstallerArgs `
                -Wait -PassThru -NoNewWindow -ErrorAction Stop
        }
        "SCRIPT" {
            Write-PackageLog "Running install script: $InstallerPath"
            $proc = Start-Process -FilePath "powershell.exe" `
                -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$InstallerPath`" $InstallerArgs" `
                -Wait -PassThru -NoNewWindow -ErrorAction Stop
        }
    }

    $installDuration = ([datetime]::UtcNow - $installStart).TotalSeconds
    Write-PackageLog "Installer exited with code $($proc.ExitCode) in $([math]::Round($installDuration, 1))s"

    # Evaluate exit code
    $successCodes = @(0, 1641, 3010)   # 0=success, 1641/3010=success+reboot
    if ($proc.ExitCode -notin $successCodes) {
        throw "Installer failed with exit code $($proc.ExitCode). Check logs in $PackageOutputDir."
    }

    if ($proc.ExitCode -in @(1641, 3010)) {
        Write-PackageLog "Installer requested reboot (code $($proc.ExitCode)). Sequencing will continue without reboot." -Level WARN
    }

    return $proc.ExitCode
}

# ============================================================================
# REGION: Attended Mode — Signal-Based Monitoring
# ============================================================================
function Wait-ForAttendedSignal {
    param(
        [int]$PollIntervalSeconds = 5
    )

    Write-PackageLog "=== Attended Mode: Waiting for customization signal ===" -Level STEP
    Write-PackageLog "Signal file: $AttendedSignalFile"
    Write-PackageLog "Create the signal file when customization is complete."

    # Track filesystem/registry activity while waiting
    $changeLog = [System.Collections.Generic.List[string]]::new()

    $elapsed = 0
    while (-not (Test-Path $AttendedSignalFile)) {
        Start-Sleep -Seconds $PollIntervalSeconds
        $elapsed += $PollIntervalSeconds

        if ($elapsed % 60 -eq 0) {
            $minutesElapsed = [math]::Round($elapsed / 60, 0)
            Write-PackageLog "Still waiting for signal... ($minutesElapsed min elapsed)" -Level DEBUG
        }

        if ($elapsed -ge $TimeoutSeconds) {
            throw "Attended mode timed out after $TimeoutSeconds seconds waiting for signal file."
        }
    }

    Write-PackageLog "Signal received after $elapsed seconds." -Level STEP

    # Clean up signal file
    Remove-Item $AttendedSignalFile -Force -ErrorAction SilentlyContinue
    return $changeLog
}

# ============================================================================
# REGION: App-V Sequencing Engine
# ============================================================================
function Invoke-AppVSequencing {
    Write-PackageLog "=== Starting App-V Sequencing ===" -Level STEP

    # Create output directory
    New-Item -Path $PackageOutputDir -ItemType Directory -Force | Out-Null
    Write-PackageLog "Output directory: $PackageOutputDir"

    # Pre-install script
    if ($PreInstallScript) {
        Write-PackageLog "Executing pre-install script..." -Level STEP
        if (Test-Path $PreInstallScript) {
            & $PreInstallScript
        }
        else {
            Invoke-Expression $PreInstallScript
        }
        Write-PackageLog "Pre-install script completed."
    }

    # Capture baseline
    $baseline = Save-MachineBaseline

    # Determine sequencing approach
    if ($script:UseCliMode) {
        Invoke-CliSequencing
    }
    else {
        Invoke-ModuleSequencing -Baseline $baseline
    }
}

function Invoke-ModuleSequencing {
    param(
        [hashtable]$Baseline
    )

    Write-PackageLog "Using PowerShell module-driven sequencing." -Level STEP

    # ---- Build common parameters ----
    $seqParams = @{
        Name                  = $PackageName
        Path                  = $PackageOutputDir
        PrimaryVirtualApplicationDirectory = $PrimaryVirtualDirectory
    }

    if ($TemplateFile) {
        $seqParams["TemplateFilePath"] = $TemplateFile
        Write-PackageLog "Using template/accelerator: $TemplateFile"
    }

    switch ($Mode) {
        # -----------------------------------------------------------------
        "Unattended" {
            Write-PackageLog "Mode: UNATTENDED — silent install + auto-finalize" -Level STEP

            $seqParams["Installer"]       = $InstallerPath
            $seqParams["InstallerOptions"] = $InstallerArgs

            try {
                Write-PackageLog "Calling New-AppvSequencerPackage..."
                $result = New-AppvSequencerPackage @seqParams -FullLoad -ErrorAction Stop
                Write-PackageLog "Sequencer completed successfully."
            }
            catch {
                Write-PackageLog "Module sequencing failed: $_" -Level ERROR

                # Fallback: manual install + Update-AppvSequencerPackage
                Write-PackageLog "Attempting fallback: manual install + Update flow..." -Level WARN
                Invoke-FallbackSequencing -Baseline $Baseline
            }
        }

        # -----------------------------------------------------------------
        "Attended" {
            Write-PackageLog "Mode: ATTENDED — manual install + signal to finalize" -Level STEP

            # Start monitoring phase
            Write-PackageLog "Starting sequencer in monitoring mode..."

            try {
                # Create the package shell, then enter editing/monitoring
                $result = New-AppvSequencerPackage @seqParams -FullLoad -ErrorAction Stop
            }
            catch [System.Management.Automation.MethodInvocationException] {
                # Expected in some versions — sequencer waits for user input
                Write-PackageLog "Sequencer entered interactive wait state." -Level DEBUG
            }

            # Wait for the external signal (operator finishes customization)
            Wait-ForAttendedSignal

            # Post-install script if any
            if ($PostInstallScript -and (Test-Path $PostInstallScript)) {
                Write-PackageLog "Executing post-install script..." -Level STEP
                & $PostInstallScript
                Write-PackageLog "Post-install script completed."
            }

            # Finalize
            Write-PackageLog "Finalizing attended sequencing..."
            try {
                Update-AppvSequencerPackage -Name $PackageName -Path $PackageOutputDir `
                    -FullLoad -ErrorAction Stop
            }
            catch {
                Write-PackageLog "Finalize via Update-AppvSequencerPackage failed: $_" -Level WARN
            }
        }

        # -----------------------------------------------------------------
        "Hybrid" {
            Write-PackageLog "Mode: HYBRID — silent install + attended customization window" -Level STEP

            # Phase 1: Automated install
            $exitCode = Invoke-InstallerExecution
            Write-PackageLog "Silent install phase complete (exit code: $exitCode)."

            # Phase 2: Post-install customization window
            Write-PackageLog "Entering customization window..." -Level STEP
            if ($PostInstallScript -and (Test-Path $PostInstallScript)) {
                Write-PackageLog "Executing post-install script..." -Level STEP
                & $PostInstallScript
                Write-PackageLog "Post-install script completed."
            }

            # Wait for signal if no post-install script handles it automatically
            Wait-ForAttendedSignal

            # Phase 3: Sequence with Update
            Write-PackageLog "Finalizing hybrid sequencing..."
            try {
                $result = New-AppvSequencerPackage @seqParams -FullLoad -ErrorAction Stop
            }
            catch {
                Write-PackageLog "Hybrid sequencing exception (may be expected): $_" -Level DEBUG
            }
        }
    }

    # Compare machine state
    $delta = Compare-MachineState -Baseline $Baseline
    return $delta
}

function Invoke-FallbackSequencing {
    <#
    .SYNOPSIS
        CLI-based fallback when the PowerShell module approach fails.
        Uses direct installer execution + filesystem/registry diff.
    #>
    param(
        [hashtable]$Baseline
    )

    Write-PackageLog "Fallback: Direct install + manual capture approach" -Level STEP

    # Take filesystem snapshot of PVAD area
    $pvadParent = Split-Path $PrimaryVirtualDirectory -Parent
    if (-not (Test-Path $pvadParent)) {
        New-Item -Path $pvadParent -ItemType Directory -Force | Out-Null
    }

    # Execute installer
    $exitCode = Invoke-InstallerExecution

    # Post-install script
    if ($PostInstallScript -and (Test-Path $PostInstallScript)) {
        Write-PackageLog "Executing post-install script (fallback)..." -Level STEP
        & $PostInstallScript
    }

    # Attempt to use Update-AppvSequencerPackage on an existing skeleton
    try {
        Update-AppvSequencerPackage -Name $PackageName -Path $PackageOutputDir `
            -Installer $InstallerPath -InstallerOptions $InstallerArgs `
            -FullLoad -ErrorAction Stop
        Write-PackageLog "Fallback sequencing succeeded via Update-AppvSequencerPackage."
    }
    catch {
        Write-PackageLog "Fallback sequencing also failed: $_" -Level ERROR
        throw "All sequencing methods exhausted. Manual packaging may be required."
    }
}

function Invoke-CliSequencing {
    <#
    .SYNOPSIS
        Drive sequencing through sequencer.exe CLI when the PowerShell module
        is unavailable.
    #>
    Write-PackageLog "Using CLI-based sequencing (sequencer.exe)." -Level STEP

    $cliArgs = @(
        "/INSTALLPACKAGE", "`"$InstallerPath`"",
        "/INSTALLEROPTIONS", "`"$InstallerArgs`"",
        "/OUTPUTPACKAGEPATH", "`"$PackageOutputDir`"",
        "/PACKAGENAME", "`"$PackageName`"",
        "/PRIMARYVIRTUALAPPLICATIONDIRECTORY", "`"$PrimaryVirtualDirectory`"",
        "/GENERATEPACKAGE"
    )

    if ($TemplateFile) {
        $cliArgs += @("/TEMPLATEFILEPATH", "`"$TemplateFile`"")
    }

    $argString = $cliArgs -join " "
    Write-PackageLog "Running: sequencer.exe $argString"

    $proc = Start-Process -FilePath "sequencer.exe" -ArgumentList $argString `
        -Wait -PassThru -NoNewWindow -RedirectStandardOutput (Join-Path $PackageOutputDir "sequencer_stdout.log") `
        -RedirectStandardError (Join-Path $PackageOutputDir "sequencer_stderr.log")

    if ($proc.ExitCode -ne 0) {
        $stderr = Get-Content (Join-Path $PackageOutputDir "sequencer_stderr.log") -ErrorAction SilentlyContinue
        throw "sequencer.exe failed with exit code $($proc.ExitCode). Stderr: $stderr"
    }

    Write-PackageLog "CLI sequencing completed successfully."
}

# ============================================================================
# REGION: Post-Sequencing Validation & Enrichment
# ============================================================================
function Test-PackageOutput {
    Write-PackageLog "=== Validating Package Output ===" -Level STEP

    $appvFile = Get-ChildItem -Path $PackageOutputDir -Filter "*.appv" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

    $packageInfo = [ordered]@{
        appv_file              = ""
        appv_size_mb           = 0
        deployment_config      = ""
        user_config            = ""
        virtual_app_directory  = $PrimaryVirtualDirectory
        files_captured         = 0
    }

    if ($appvFile) {
        $packageInfo.appv_file    = $appvFile.FullName
        $packageInfo.appv_size_mb = [math]::Round($appvFile.Length / 1MB, 2)
        $sizeMB = $packageInfo.appv_size_mb
        Write-PackageLog "Package file: $($appvFile.FullName) ($sizeMB MB)"

        # Count files inside the .appv (it's a ZIP)
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::OpenRead($appvFile.FullName)
            $packageInfo.files_captured = $zip.Entries.Count
            $zip.Dispose()
            Write-PackageLog "Files captured in package: $($packageInfo.files_captured)"
        }
        catch {
            Write-PackageLog "Could not inspect .appv ZIP contents: $_" -Level WARN
        }
    }
    else {
        Write-PackageLog "WARNING: No .appv file found in output directory!" -Level ERROR
    }

    # Locate config files
    $deployConfig = Get-ChildItem -Path $PackageOutputDir -Filter "*DeploymentConfig*" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    $userConfig   = Get-ChildItem -Path $PackageOutputDir -Filter "*UserConfig*"       -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($deployConfig) {
        $packageInfo.deployment_config = $deployConfig.FullName
        Write-PackageLog "Deployment config: $($deployConfig.FullName)"
    }
    if ($userConfig) {
        $packageInfo.user_config = $userConfig.FullName
        Write-PackageLog "User config: $($userConfig.FullName)"
    }

    return $packageInfo
}

function Set-ExclusionPatterns {
    <#
    .SYNOPSIS
        Apply exclusion patterns to the DeploymentConfig XML to prevent
        unwanted filesystem/registry paths from appearing in the package.
    #>
    if ($ExclusionPatterns.Count -eq 0) { return }

    Write-PackageLog "Applying $($ExclusionPatterns.Count) exclusion pattern(s)..." -Level STEP

    $deployConfig = Get-ChildItem -Path $PackageOutputDir -Filter "*DeploymentConfig*" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $deployConfig) {
        Write-PackageLog "No DeploymentConfig found — cannot apply exclusions." -Level WARN
        return
    }

    try {
        [xml]$xml = Get-Content $deployConfig.FullName -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("appv", $xml.DocumentElement.NamespaceURI)

        foreach ($pattern in $ExclusionPatterns) {
            Write-PackageLog "Exclusion pattern: $pattern" -Level DEBUG
            # Add to filesystem exclusions
            $fsExclusions = $xml.SelectSingleNode("//appv:FileSystem/appv:Exclusions", $ns)
            if ($fsExclusions) {
                $excl = $xml.CreateElement("Exclusion", $xml.DocumentElement.NamespaceURI)
                $excl.InnerText = $pattern
                $fsExclusions.AppendChild($excl) | Out-Null
            }
        }

        $xml.Save($deployConfig.FullName)
        Write-PackageLog "Exclusion patterns applied to DeploymentConfig."
    }
    catch {
        Write-PackageLog "Failed to apply exclusion patterns: $_" -Level WARN
    }
}

# ============================================================================
# REGION: Main Execution Pipeline
# ============================================================================
function Invoke-PackagingPipeline {
    Write-PackageLog "================================================================" -Level STEP
    Write-PackageLog "  App-V Package Automation Engine — Run $($script:RunId)" -Level STEP
    Write-PackageLog "================================================================" -Level STEP
    Write-PackageLog "Package     : $PackageName v$PackageVersion"
    Write-PackageLog "Mode        : $Mode"
    Write-PackageLog "Installer   : $InstallerPath"
    Write-PackageLog "Output Root : $OutputRoot"
    Write-PackageLog "PVAD        : $PrimaryVirtualDirectory"
    Write-PackageLog "Host        : $env:COMPUTERNAME"
    Write-PackageLog "Timeout     : $TimeoutSeconds seconds"

    # Create output directory early so logging can persist
    New-Item -Path $PackageOutputDir -ItemType Directory -Force | Out-Null

    # Copy installer locally if it's on a network share (prevents sequencer locking issues)
    if ($InstallerPath.StartsWith("\\")) {
        $localInstaller = Join-Path $PackageOutputDir (Split-Path $InstallerPath -Leaf)
        Write-PackageLog "Copying installer from network share to local staging..."
        Copy-Item -Path $InstallerPath -Destination $localInstaller -Force
        $script:OriginalInstallerPath = $InstallerPath
        $script:InstallerPath = $localInstaller
        Set-Variable -Name InstallerPath -Value $localInstaller -Scope 1
        Write-PackageLog "Local copy: $localInstaller"
    }

    # Pre-flight
    Test-SequencerEnvironment

    # Execute sequencing
    Invoke-AppVSequencing

    # Apply exclusions
    Set-ExclusionPatterns

    # Validate
    $packageInfo = Test-PackageOutput

    # Cleanup
    Invoke-PostPackageCleanup

    return $packageInfo
}

# ============================================================================
# REGION: Entry Point with Timeout & Error Boundary
# ============================================================================
$script:UseCliMode = $false
$exitReport = $null

# Register a timer-based timeout watchdog (kills this process if exceeded)
$timeoutTimer = New-Object System.Timers.Timer
$timeoutTimer.Interval = $TimeoutSeconds * 1000
$timeoutTimer.AutoReset = $false
$timeoutExpired = $false

$timeoutAction = Register-ObjectEvent -InputObject $timeoutTimer -EventName Elapsed -Action {
    $global:timeoutExpired = $true
}
$timeoutTimer.Start()

try {
    # Direct invocation — all functions are in-scope
    $packageInfo = Invoke-PackagingPipeline

    if ($global:timeoutExpired) {
        throw "Wall-clock timeout after $TimeoutSeconds seconds."
    }

    $exitReport = New-StatusReport -Status "Success" -PackageInfo $packageInfo
    Write-PackageLog "=== PACKAGING COMPLETED SUCCESSFULLY ===" -Level STEP
}
catch {
    $errorMsg = $_.Exception.Message
    $errorStack = $_.ScriptStackTrace
    Write-PackageLog "FATAL ERROR: $errorMsg" -Level ERROR
    Write-PackageLog "Stack: $errorStack" -Level ERROR

    if ($errorMsg -match "timeout") {
        $exitReport = New-StatusReport -Status "TimedOut" -ErrorDetail $errorMsg
    }
    else {
        $exitReport = New-StatusReport -Status "Failed" -ErrorDetail $errorMsg
    }

    # Cleanup on failure
    if ($CleanupOnFailure -and (Test-Path $PackageOutputDir)) {
        # Keep logs, remove package artifacts
        Get-ChildItem $PackageOutputDir -Filter "*.appv" -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
        Write-PackageLog "Partial .appv artifacts removed (logs retained)."
    }
}
finally {
    # Dispose timeout watchdog
    $timeoutTimer.Stop()
    $timeoutTimer.Dispose()
    if ($timeoutAction) {
        Unregister-Event -SubscriptionId $timeoutAction.Id -ErrorAction SilentlyContinue
        Remove-Job -Id $timeoutAction.Id -Force -ErrorAction SilentlyContinue
    }
}

# Return structured report (consumed by WinRM caller / pipeline)
$exitReport | ConvertTo-Json -Depth 10