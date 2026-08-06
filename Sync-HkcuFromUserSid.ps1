# Cloudpaging Studio - HKCU Registry Sync
#
# Copyright (c) 2023-2026 Numecent, Inc.
#
# See the LICENSE file in the repository root for license terms.

<#
.SYNOPSIS
    Replicates registry changes an installer wrote under a specific user's HKEY_USERS\<SID> hive into this
    process's own live HKEY_CURRENT_USER, so Cloudpaging Studio's capture picks them up.
.DESCRIPTION
    When packaging is driven remotely (e.g. over WinRM) with no user interactively logged on, an installer's
    per-user registry writes commonly land under HKEY_USERS\<SID> of whichever profile happens to be loaded
    (or a profile explicitly loaded for this purpose) rather than a live HKEY_CURRENT_USER - because there is
    no interactive session to map HKCU to. Studio's capture, however, is watching this process's own session,
    so those writes are never observed.

    This script works around that by taking two snapshots around the installer run:
      Before - export the target SID's chosen subkeys (default: Software) before the installer runs.
      After  - export the same subkeys again after the installer runs, diff against the Before snapshot to
               find only what the installer actually added/changed, rewrite that delta from
               "HKEY_USERS\<SID>\..." to "HKEY_CURRENT_USER\...", and import it with regedit /s into this
               process's own live HKCU.

    Only additions/changes are synced; values the installer removed are not tracked, since the goal is
    reproducing what the installer added, not full hive mirroring.

    If the target SID's hive is not already loaded (no live session for that user), this script loads it
    on demand from that profile's NTUSER.DAT (via the ProfileList registry) and unloads it again after the
    "After" pass - but only if this script is the one that loaded it; an already-loaded (live) hive is left
    alone.

    Intended to be invoked twice around the installer command, from the same capture session so the import in
    the "After" pass lands in the same HKCU that Cloudpaging Studio is watching:
      studio-nip.ps1 wires this in automatically via the CaptureSettings.HkcuRegistrySync JSON section
      (Before runs via PreCaptureCommands, After is appended to the end of Installer.bat).
.PARAMETER Mode
    Before or After.
.PARAMETER Sid
    SID of the user whose HKEY_USERS hive should be watched (e.g. S-1-5-21-...-1001).
.PARAMETER SnapshotFolder
    Working folder for the before/after/merge .reg files and the hive-loaded marker; must be the same path
    for both the Before and After calls for a given packaging run.
.PARAMETER SubKeys
    One or more subkey paths (relative to the hive root) to watch, e.g. "Software" or "Software\MyVendor".
    Defaults to "Software", which covers the large majority of per-user application settings.
.PARAMETER KeepSnapshots
    When set, the before/after/merge .reg files in SnapshotFolder are kept after the After pass instead of
    being deleted, useful for troubleshooting what was/wasn't synced.

.EXAMPLE
    >Sync-HkcuFromUserSid.ps1 -Mode Before -Sid 'S-1-5-21-111-222-333-1001' -SnapshotFolder 'C:\NIP_software\MyApp\_HkcuSync'
    ... installer runs here ...
    >Sync-HkcuFromUserSid.ps1 -Mode After -Sid 'S-1-5-21-111-222-333-1001' -SnapshotFolder 'C:\NIP_software\MyApp\_HkcuSync'

.NOTES
    Requires Administrator rights (to load/unload registry hives) and reg.exe/regedit.exe, both built into
    Windows. Run as the same user/session that Cloudpaging Studio's capture is watching.
#>

param(
    [Parameter(Mandatory = $true)][ValidateSet('Before', 'After')]
    [string]$Mode,

    [Parameter(Mandatory = $true)][ValidatePattern('^S-1-(5-21-\d+-\d+-\d+-\d+|5-\d+)$')]
    [string]$Sid,

    [Parameter(Mandatory = $true)]
    [string]$SnapshotFolder,

    [Parameter(Mandatory = $false)]
    [string[]]$SubKeys = @('Software'),

    [Parameter(Mandatory = $false)]
    [switch]$KeepSnapshots
)

$ErrorActionPreference = 'Stop'

function Test-HiveLoaded {
    param([Parameter(Mandatory = $true)][string]$Sid)
    return Test-Path -Path "Registry::HKEY_USERS\$Sid"
}

function ConvertTo-SafeFileName {
    param([Parameter(Mandatory = $true)][string]$Text)
    return ($Text -replace '[\\/:*?"<>|]', '_')
}

function Mount-UserHiveIfNeeded {
    param(
        [Parameter(Mandatory = $true)][string]$Sid,
        [Parameter(Mandatory = $true)][string]$MarkerPath
    )

    if (Test-HiveLoaded -Sid $Sid) {
        Write-Output "HKCU sync: HKEY_USERS\$Sid is already loaded (live session); will not unload it afterward."
        return
    }

    $profileKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Sid"
    if (-NOT (Test-Path -Path $profileKey)) {
        Throw "SID '$Sid' has no loaded hive and no entry under ProfileList; cannot locate its NTUSER.DAT. Verify the SID is correct."
    }

    $profilePath = (Get-ItemProperty -Path $profileKey -Name ProfileImagePath -ErrorAction Stop).ProfileImagePath
    $ntUserDat = Join-Path $profilePath "NTUSER.DAT"
    if (-NOT (Test-Path -Path $ntUserDat)) {
        Throw "NTUSER.DAT not found at '$ntUserDat' for SID '$Sid'."
    }

    $loadOutput = & reg.exe load "HKU\$Sid" "$ntUserDat" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Throw "Failed to load hive for SID '$Sid' from '$ntUserDat': $loadOutput"
    }

    Set-Content -Path $MarkerPath -Value "loaded by Sync-HkcuFromUserSid.ps1 at $(Get-Date -Format o)" -Force
    Write-Output "HKCU sync: loaded HKEY_USERS\$Sid from $ntUserDat"
}

function Dismount-UserHiveIfWeLoadedIt {
    param(
        [Parameter(Mandatory = $true)][string]$Sid,
        [Parameter(Mandatory = $true)][string]$MarkerPath
    )

    if (-NOT (Test-Path -Path $MarkerPath)) { return }

    # Release any lingering registry provider handles this process holds under HKU\<Sid> before unloading
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500

    $attempt = 0
    $unloaded = $false
    do {
        $attempt++
        $unloadOutput = & reg.exe unload "HKU\$Sid" 2>&1
        $unloaded = ($LASTEXITCODE -eq 0)
        if (-NOT $unloaded -and $attempt -lt 5) {
            Start-Sleep -Milliseconds 1000
            [System.GC]::Collect()
        }
    } while (-NOT $unloaded -and $attempt -lt 5)

    if ($unloaded) {
        Remove-Item -Path $MarkerPath -Force -ErrorAction SilentlyContinue
        Write-Output "HKCU sync: unloaded HKEY_USERS\$Sid"
    }
    else {
        Write-Warning "HKCU sync: could not unload HKEY_USERS\$Sid after $attempt attempt(s): $unloadOutput. It may stay loaded until reboot or a manual 'reg unload HKU\$Sid'."
    }
}

function Export-UserHiveSubKeys {
    param(
        [Parameter(Mandatory = $true)][string]$Sid,
        [Parameter(Mandatory = $true)][string[]]$SubKeys,
        [Parameter(Mandatory = $true)][string]$DestFolder,
        [Parameter(Mandatory = $true)][string]$Suffix
    )

    $exported = @{}
    foreach ($sub in $SubKeys) {
        $destFile = Join-Path $DestFolder "$(ConvertTo-SafeFileName $sub).$Suffix.reg"
        # Exit code 1 ("key not found") just means nothing exists there yet - an empty baseline is fine
        & reg.exe export "HKU\$Sid\$sub" $destFile /y 2>$null | Out-Null
        $exported[$sub] = $destFile
    }
    return $exported
}

function ConvertFrom-RegExport {
    param([Parameter(Mandatory = $true)][string]$Path)

    $keys = [ordered]@{}
    if (-NOT (Test-Path -Path $Path)) { return $keys }

    $content = [System.IO.File]::ReadAllText($Path)
    if (-NOT $content) { return $keys }

    $rawLines = $content -split "`r`n|`n"

    # Join hex: value continuation lines (a trailing backslash means the value continues on the next line)
    $lines = New-Object System.Collections.Generic.List[string]
    $buffer = $null
    foreach ($line in $rawLines) {
        $buffer = if ($null -ne $buffer) { "$buffer`n$line" } else { $line }
        if ($buffer.TrimEnd() -match '\\$') { continue }
        $lines.Add($buffer)
        $buffer = $null
    }
    if ($buffer) { $lines.Add($buffer) }

    $currentKey = $null
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if (-NOT $trimmed -or $trimmed.StartsWith(';') -or $trimmed -like 'Windows Registry Editor*') { continue }

        if ($trimmed.StartsWith('[') -and $trimmed.EndsWith(']')) {
            $currentKey = $trimmed.Substring(1, $trimmed.Length - 2)
            if (-NOT $keys.Contains($currentKey)) { $keys[$currentKey] = [ordered]@{} }
            continue
        }

        if (-NOT $currentKey) { continue }

        $eqIndex = $line.IndexOf('=')
        if ($eqIndex -lt 0) { continue }

        $namePart = $line.Substring(0, $eqIndex).Trim()
        $valuePart = $line.Substring($eqIndex + 1)
        $keys[$currentKey][$namePart] = $valuePart
    }

    return $keys
}

function Compare-RegSnapshots {
    param(
        [Parameter(Mandatory = $true)]$Before,
        [Parameter(Mandatory = $true)]$After
    )

    $delta = [ordered]@{}

    foreach ($key in $After.Keys) {
        $beforeValues = if ($Before.Contains($key)) { $Before[$key] } else { $null }
        $afterValues = $After[$key]

        $changed = [ordered]@{}
        foreach ($name in $afterValues.Keys) {
            $newVal = $afterValues[$name]
            $oldVal = if ($beforeValues -and $beforeValues.Contains($name)) { $beforeValues[$name] } else { $null }
            if ($oldVal -ne $newVal) {
                $changed[$name] = $newVal
            }
        }

        if ($changed.Count -gt 0) {
            $delta[$key] = $changed
        }
    }

    return $delta
}

function ConvertTo-RegExportText {
    param(
        [Parameter(Mandatory = $true)]$Keys,
        [Parameter(Mandatory = $true)][string]$SourceHivePrefix,
        [Parameter(Mandatory = $true)][string]$TargetHivePrefix
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Windows Registry Editor Version 5.00")
    [void]$sb.AppendLine()

    foreach ($key in $Keys.Keys) {
        $targetKey = $key
        if ($targetKey.StartsWith($SourceHivePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $targetKey = $TargetHivePrefix + $targetKey.Substring($SourceHivePrefix.Length)
        }
        [void]$sb.AppendLine("[$targetKey]")
        foreach ($name in $Keys[$key].Keys) {
            [void]$sb.AppendLine("$name=$($Keys[$key][$name])")
        }
        [void]$sb.AppendLine()
    }

    return $sb.ToString()
}

$markerPath = Join-Path $SnapshotFolder "$(ConvertTo-SafeFileName $Sid).hiveloaded.marker"

switch ($Mode) {
    "Before" {
        if (Test-Path -Path $SnapshotFolder) {
            Remove-Item -Path $SnapshotFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $SnapshotFolder -Force | Out-Null

        Mount-UserHiveIfNeeded -Sid $Sid -MarkerPath $markerPath
        [void](Export-UserHiveSubKeys -Sid $Sid -SubKeys $SubKeys -DestFolder $SnapshotFolder -Suffix "before")

        Write-Output "HKCU sync: captured 'before' snapshot for HKEY_USERS\$Sid ($($SubKeys -join ', '))"
    }

    "After" {
        if (-NOT (Test-Path -Path $SnapshotFolder)) {
            Throw "SnapshotFolder '$SnapshotFolder' does not exist; -Mode Before must run first with the same -SnapshotFolder."
        }

        if (-NOT (Test-HiveLoaded -Sid $Sid)) {
            Mount-UserHiveIfNeeded -Sid $Sid -MarkerPath $markerPath
        }

        $afterFiles = Export-UserHiveSubKeys -Sid $Sid -SubKeys $SubKeys -DestFolder $SnapshotFolder -Suffix "after"

        $totalKeys = 0
        $totalValues = 0
        $sourcePrefix = "HKEY_USERS\$Sid"

        foreach ($sub in $SubKeys) {
            $beforeFile = Join-Path $SnapshotFolder "$(ConvertTo-SafeFileName $sub).before.reg"
            $afterFile = $afterFiles[$sub]

            $beforeKeys = ConvertFrom-RegExport -Path $beforeFile
            $afterKeys = ConvertFrom-RegExport -Path $afterFile

            $delta = Compare-RegSnapshots -Before $beforeKeys -After $afterKeys
            if ($delta.Count -eq 0) { continue }

            $mergeText = ConvertTo-RegExportText -Keys $delta -SourceHivePrefix $sourcePrefix -TargetHivePrefix "HKEY_CURRENT_USER"
            $mergeFile = Join-Path $SnapshotFolder "$(ConvertTo-SafeFileName $sub).merge.reg"
            $mergeText | Out-File -FilePath $mergeFile -Encoding unicode

            & regedit.exe /s $mergeFile
            Start-Sleep -Milliseconds 300

            $totalKeys += $delta.Count
            foreach ($k in $delta.Keys) { $totalValues += $delta[$k].Count }
        }

        Write-Output "HKCU sync: synced $totalKeys key(s) / $totalValues value(s) from HKEY_USERS\$Sid into the live HKEY_CURRENT_USER."

        Dismount-UserHiveIfWeLoadedIt -Sid $Sid -MarkerPath $markerPath

        if (-NOT $KeepSnapshots) {
            Remove-Item -Path $SnapshotFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
