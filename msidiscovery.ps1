# Cloudpaging Automated-Packaging - MSI Discovery Report
#
# Copyright (c) 2023-2026 Numecent, Inc.
#
# See the LICENSE file in the repository root for license terms.

<#
.SYNOPSIS
    Reads a Windows Installer (.msi) database directly - without installing or running it - to discover
    add-ins, COM registrations, drivers, files, registry changes, and other packaging-relevant metadata.
.DESCRIPTION
    Uses the Windows Installer Automation COM object (WindowsInstaller.Installer) to open the MSI database
    read-only and query its internal tables. This is the same engine msiexec itself uses, so it is a far more
    reliable source of truth than string-scanning the binary or extracting the MSI with a third-party tool.

    The report identifies:
      - COM add-ins: registry rows matching the Office "...\Office\<App>\Addins\<ProgId>" pattern, cross
        referenced against the ProgId/Class/Component/File tables to resolve the CLSID and backing binary.
      - Excel classic add-ins (XLL/XLA/XLAM): registry rows under "...\Office\<Ver>\Excel\Options" with a
        value name of OPEN, OPEN1, OPEN2, etc. The actual value (the load string/path Excel runs at startup)
        is extracted directly, since that is the value that matters for repackaging.
      - Add-in registration performed outside the declarative Registry table: many installers (WiX/InstallShield
        custom setups) write the "OPEN" key or the COM Addins keys from a Custom Action instead - a VBScript/
        JScript or a compiled DLL/EXE stored as an opaque stream in the "Binary" table. Since that code isn't
        declarative, it can't be resolved to a guaranteed final value; instead every Binary table stream is
        extracted (via Database.Export, never executed) and scanned for literal strings that indicate Excel/
        Office add-in registration (registry path fragments, the literal "OPEN"/"OPENn" value name, .xll/.xla/
        .xlam paths, MSI "[PROPERTY]" formatted-string placeholders). Each CustomAction's Target/Source text is
        scanned the same way. These are reported as "likely" hints, not confirmed values, and are clearly
        labeled as such. Installed .xll/.xla/.xlam files are also cross-referenced from the File table as an
        independent, always-reliable signal that an Excel add-in is present even when its load mechanism can't
        be statically resolved.
      - Drivers: ServiceInstall rows whose ServiceType flags mark them as a kernel or file-system driver,
        any File table entries with a .sys extension, and the ODBCDriver table when present.
      - HKEY_CURRENT_USER registry hives: every distinct registry key written under HKCU (including Root=-1
        rows resolved via the ALLUSERS property), with the value names present but not their data.
      - Full file, registry, component, and directory inventories, with a best-effort (approximate) resolved
        install path for each file.
      - Summary Information stream, Property table, Feature tree, Custom Actions, Shortcuts, Environment,
        IniFile, and Upgrade table entries, plus a full list of every table present in the database so nothing
        vendor-specific gets silently missed.

    Because this only opens the database in read-only mode, it never modifies the MSI, never requires
    administrator rights, and never executes any installer or custom action code - including the binary/script
    streams pulled from the Binary table, which are only extracted to a temp folder and string-scanned, never run.
.PARAMETER MsiPath
    Path to one or more .msi files to scan. Accepts pipeline input, so it can be combined with Get-ChildItem
    to scan every MSI under a folder.
.PARAMETER OutputJsonPath
    Optional path to save the full report as JSON. When scanning multiple files via the pipeline, "_<n>" is
    appended before the extension for each additional file.
.PARAMETER Quiet
    Suppresses the human-readable console summary; the full report object is still returned on the pipeline.
.PARAMETER SkipBinaryStreamScan
    Skips extracting and string-scanning the Binary table (custom action scripts/DLLs/EXEs). Use this for a
    faster pass when you only care about the declarative Registry table, or on MSIs with very large embedded
    binaries.
.PARAMETER MaxBinaryStreamScanBytes
    Per-stream size cap for the Binary table scan; streams larger than this are skipped (and reported as
    skipped) rather than fully scanned, to keep large embedded payloads from stalling the scan. Default 20MB.

.EXAMPLE
    >Get-MsiDiscoveryReport.ps1 -MsiPath 'C:\NIP_software\MyAddin\Installer_Cfg\MyAddin.msi' -OutputJsonPath 'C:\NIP_software\MyAddin\discovery.json'

.EXAMPLE
    >Get-ChildItem 'C:\NIP_software' -Recurse -Filter *.msi | Get-MsiDiscoveryReport.ps1 | Where-Object { $_.AddIns.ComAddIns.Count -gt 0 -or $_.AddIns.ExcelOpenKeyAddIns.Count -gt 0 }

.NOTES
    Requires Windows Installer (present on all supported Windows 10/11 systems); no external modules needed.
#>

param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('FullName')]
    [string[]]$MsiPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputJsonPath,

    [Parameter(Mandatory = $false)]
    [switch]$Quiet,

    [Parameter(Mandatory = $false)]
    [switch]$SkipBinaryStreamScan,

    [Parameter(Mandatory = $false)]
    [long]$MaxBinaryStreamScanBytes = 20MB
)

begin {
    $ErrorActionPreference = 'Stop'

    # msiOpenDatabaseModeReadOnly - never opens the MSI for write, so the source file cannot be modified
    $MSI_OPEN_READONLY = 0

    $script:SpecialFolders = @{
        "TARGETDIR"             = "C:\"
        "SourceDir"             = "C:\"
        "ProgramFilesFolder"    = "C:\Program Files\"
        "ProgramFiles64Folder"  = "C:\Program Files\"
        "CommonFilesFolder"     = "C:\Program Files\Common Files\"
        "CommonFiles64Folder"   = "C:\Program Files\Common Files\"
        "SystemFolder"          = "C:\Windows\System32\"
        "System16Folder"        = "C:\Windows\System\"
        "System64Folder"        = "C:\Windows\System32\"
        "WindowsFolder"         = "C:\Windows\"
        "WindowsVolume"         = "C:\"
        "AppDataFolder"         = "%AppData%\"
        "LocalAppDataFolder"    = "%LocalAppData%\"
        "PersonalFolder"        = "%UserProfile%\Documents\"
        "DesktopFolder"         = "%UserProfile%\Desktop\"
        "FavoritesFolder"       = "%UserProfile%\Favorites\"
        "SendToFolder"          = "%AppData%\Microsoft\Windows\SendTo\"
        "StartupFolder"         = "%AppData%\Microsoft\Windows\Start Menu\Programs\Startup\"
        "StartMenuFolder"       = "%AppData%\Microsoft\Windows\Start Menu\"
        "ProgramMenuFolder"     = "%AppData%\Microsoft\Windows\Start Menu\Programs\"
        "CommonAppDataFolder"   = "%ProgramData%\"
        "TempFolder"            = "%Temp%\"
        "FontsFolder"           = "C:\Windows\Fonts\"
    }

    $RegistryRootMap = @{
        "-1" = "HKEY_CURRENT_USER or HKEY_LOCAL_MACHINE (runtime dependent on ALLUSERS)"
        "0"  = "HKEY_CLASSES_ROOT"
        "1"  = "HKEY_CURRENT_USER"
        "2"  = "HKEY_LOCAL_MACHINE"
        "3"  = "HKEY_USERS"
    }

    # The Record object returned by View.ColumnInfo() does not reliably expose FieldCount through
    # PowerShell's late-bound COM interop (it comes back empty even though StringData(i) works fine).
    # Column names/order are instead read from the "_Columns" system table, which is authoritative and
    # works through plain row fetches. A few system tables (leading underscore) do not describe
    # themselves in "_Columns", so their column lists are hardcoded here.
    $SystemTableColumns = @{
        "_Tables"  = @("Name")
        "_Columns" = @("Table", "Number", "Name", "Type")
        "_Streams" = @("Name", "Data")
        "_Storages"= @("Name", "Data")
    }

    function Test-MsiTableExists {
        param(
            [Parameter(Mandatory = $true)]$Database,
            [Parameter(Mandatory = $true)][string]$TableName
        )

        # Meta tables never list themselves as a row inside "_Tables", but they always exist
        if ($SystemTableColumns.ContainsKey($TableName)) {
            return $true
        }

        $view = $Database.OpenView("SELECT ``Name`` FROM ``_Tables`` WHERE ``Name`` = '$TableName'")
        [void]$view.Execute()
        $record = $view.Fetch()
        $exists = $null -ne $record
        if ($record) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
        [void]$view.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view)
        return $exists
    }

    function Get-MsiTableColumnNames {
        param(
            [Parameter(Mandatory = $true)]$Database,
            [Parameter(Mandatory = $true)][string]$TableName
        )

        if ($SystemTableColumns.ContainsKey($TableName)) {
            return $SystemTableColumns[$TableName]
        }

        $view = $Database.OpenView("SELECT ``Name`` FROM ``_Columns`` WHERE ``Table`` = '$TableName' ORDER BY ``Number``")
        [void]$view.Execute()

        $names = @()
        while ($true) {
            $record = $view.Fetch()
            if (-NOT $record) { break }
            $names += $record.StringData(1)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record)
        }

        [void]$view.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view)
        return $names
    }

    function Get-MsiTableRows {
        param(
            [Parameter(Mandatory = $true)]$Database,
            [Parameter(Mandatory = $true)][string]$TableName
        )

        if (-NOT (Test-MsiTableExists -Database $Database -TableName $TableName)) {
            return @()
        }

        $names = @(Get-MsiTableColumnNames -Database $Database -TableName $TableName)
        if ($names.Count -eq 0) {
            Write-Warning "Could not determine columns for table '$TableName'; skipping it."
            return @()
        }
        $fieldCount = $names.Count

        $view = $Database.OpenView("SELECT * FROM ``$TableName``")
        [void]$view.Execute()

        $rows = @()
        while ($true) {
            $record = $view.Fetch()
            if (-NOT $record) { break }

            $row = [ordered]@{}
            for ($i = 1; $i -le $fieldCount; $i++) {
                $row[$names[$i - 1]] = $record.StringData($i)
            }
            $rows += [PSCustomObject]$row
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record)
        }

        [void]$view.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view)
        return $rows
    }

    function Get-MsiSummaryInfo {
        param([Parameter(Mandatory = $true)]$Database)

        $summaryInfo = $Database.SummaryInformation(0)

        function Get-SummaryProperty {
            param($SummaryInfo, [int]$PropertyId)
            try {
                $value = $SummaryInfo.Property($PropertyId)
                if ($value -eq "") { return $null }
                return $value
            }
            catch {
                return $null
            }
        }

        $wordCountRaw = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 15
        $wordCountInt = 0
        [void][int]::TryParse($wordCountRaw, [ref]$wordCountInt)

        $securityRaw = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 19
        $securityText = switch ($securityRaw) {
            "0" { "No restriction (readable/writable)" }
            "2" { "Read-only recommended" }
            "4" { "Read-only enforced" }
            default { "Unknown" }
        }

        $result = [PSCustomObject]@{
            Title                 = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 2
            Subject               = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 3
            Author                = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 4
            Keywords              = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 5
            Comments              = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 6
            PlatformAndLanguage   = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 7   # Template: "Platform;LanguageID"
            LastSavedBy           = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 8
            PackageCode           = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 9   # RevisionNumber
            CreatedDate           = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 12
            LastSavedDate         = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 13
            MinimumInstallerEngine= Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 14  # PageCount, e.g. 200 = MSI 2.0
            SourceFileMode        = [PSCustomObject]@{
                Raw               = $wordCountRaw
                ShortFileNamesOnly= (($wordCountInt -band 0x1) -ne 0)
                Compressed        = (($wordCountInt -band 0x2) -ne 0)
                AdminImage        = (($wordCountInt -band 0x4) -ne 0)
                LimitedUserPatch  = (($wordCountInt -band 0x8) -ne 0)
            }
            ApplicationName       = Get-SummaryProperty -SummaryInfo $summaryInfo -PropertyId 18
            Security              = $securityText
        }

        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($summaryInfo)
        return $result
    }

    function Get-MsiPropertiesTable {
        param([Parameter(Mandatory = $true)]$Database)

        $rows = @(Get-MsiTableRows -Database $Database -TableName "Property")
        $props = [ordered]@{}
        foreach ($row in $rows) {
            $props[$row.Property] = $row.Value
        }
        return [PSCustomObject]$props
    }

    function Get-MsiDefaultDirLongName {
        param([string]$DefaultDir)

        if (-NOT $DefaultDir) { return "" }
        $targetPart = $DefaultDir.Split(':')[0]
        $parts = $targetPart.Split('|')
        if ($parts.Count -ge 2) { return $parts[1] } else { return $parts[0] }
    }

    # --- Property-driven directory resolution -----------------------------------------------------------
    #
    # The Directory table's DefaultDir is only a fallback. Real Windows Installer resolves a directory by
    # checking FIRST whether a Property exists with the exact same name as the Directory table key - if so,
    # that Property's (fully formatted) value wins outright. Installers built with Advanced Installer/WiX/
    # InstallShield routinely compute the true install directory this way via "SetProperty"-style custom
    # actions (msidbCustomActionTypeTextData, base Type value 3 regardless of the upper source/timing bits -
    # e.g. WiX's <SetProperty> element always compiles to exactly this), such as:
    #   SET_APPDIR (Type 307): Source=APPDIR, Target=[ProgramFilesFolder][Manufacturer]\[ProductShortName]
    #   SET_TARGETDIR_TO_APPDIR (Type 51): Source=TARGETDIR, Target=[APPDIR]
    # Reproducing this is the difference between reporting the meaningless literal Directory-table default
    # ("C:\APPDIR\") and the actual folder the app installs into ("C:\Program Files\Vendor\Product\").

    function Resolve-MsiFormattedString {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value,
            [Parameter(Mandatory = $true)][hashtable]$Properties
        )

        if (-NOT $Value) { return $Value }

        $evaluator = {
            param($m)
            $propName = $m.Groups[1].Value
            if ($Properties.ContainsKey($propName) -and $Properties[$propName]) {
                return $Properties[$propName]
            }
            return $m.Value
        }.GetNewClosure()

        return [regex]::Replace($Value, '\[([A-Za-z_][A-Za-z0-9_\.]*)\]', $evaluator)
    }

    function Get-MsiResolvedProperties {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$PropertyRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$CustomActionRows
        )

        $resolved = @{}
        # Seed with generic special-folder defaults (best-effort local-machine guesses)...
        # Keep the trailing backslash: standard MSI folder properties always end with one, and formatted
        # strings like "[ProgramFilesFolder][Manufacturer]\[ProductShortName]" rely on it being there.
        foreach ($k in $script:SpecialFolders.Keys) { $resolved[$k] = $script:SpecialFolders[$k] }
        # ...then the MSI's own Property table, which wins over the generic guesses when both exist
        foreach ($row in $PropertyRows) {
            if ($row.Property) { $resolved[$row.Property] = $row.Value }
        }

        # "SetProperty"-style custom actions: Source = property name being set, Target = formatted value.
        # This is the base Type-3 (TextData) family regardless of the upper source/timing bits, which are
        # not being decoded here since only the "does this action set a property" fact matters for this.
        $setPropertyActions = @($CustomActionRows | Where-Object {
                $t = 0
                [void][int]::TryParse($_.Type, [ref]$t)
                (($t -band 0x7) -eq 3) -and $_.Source
            })

        # Multiple passes resolve chained references (e.g. TARGETDIR=[APPDIR], APPDIR=[ProgramFilesFolder]...)
        # regardless of what order the custom actions happen to appear in the table.
        for ($pass = 0; $pass -lt 5; $pass++) {
            foreach ($ca in $setPropertyActions) {
                $resolved[$ca.Source] = Resolve-MsiFormattedString -Value $ca.Target -Properties $resolved
            }
        }

        return $resolved
    }

    function Resolve-MsiDirectoryPath {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][hashtable]$DirectoryLookup,
            [Parameter(Mandatory = $true)][string]$DirectoryKey,
            [Parameter(Mandatory = $true)][System.Collections.Generic.Dictionary[string, string]]$Cache,
            [Parameter(Mandatory = $true)][hashtable]$ResolvedProperties
        )

        if ($Cache.ContainsKey($DirectoryKey)) { return $Cache[$DirectoryKey] }

        # A Property with the same name as this Directory key overrides the Directory table's default -
        # only trust it if it fully resolved (no leftover "[...]" token for an unknown/unresolved property).
        if ($ResolvedProperties.ContainsKey($DirectoryKey) -and $ResolvedProperties[$DirectoryKey] -and ($ResolvedProperties[$DirectoryKey] -notmatch '\[')) {
            $resolved = $ResolvedProperties[$DirectoryKey].TrimEnd('\') + '\'
            $Cache[$DirectoryKey] = $resolved
            return $resolved
        }

        if (-NOT $DirectoryLookup.ContainsKey($DirectoryKey)) {
            # Unknown directory reference (e.g. a custom-action defined folder); flag it rather than guess
            $resolved = "[Unresolved:$DirectoryKey]\"
            $Cache[$DirectoryKey] = $resolved
            return $resolved
        }

        $entry = $DirectoryLookup[$DirectoryKey]
        $longName = Get-MsiDefaultDirLongName -DefaultDir $entry.DefaultDir

        if (-NOT $entry.Parent -or $entry.Parent -eq $DirectoryKey) {
            $targetRoot = if ($ResolvedProperties.ContainsKey('TARGETDIR')) { $ResolvedProperties['TARGETDIR'].TrimEnd('\') + '\' } else { $script:SpecialFolders['TARGETDIR'] }
            $resolved = "$targetRoot$longName\"
        }
        else {
            $parentResolved = Resolve-MsiDirectoryPath -DirectoryLookup $DirectoryLookup -DirectoryKey $entry.Parent -Cache $Cache -ResolvedProperties $ResolvedProperties
            if ($longName -and $longName -ne ".") {
                $resolved = "$parentResolved$longName\"
            }
            else {
                $resolved = $parentResolved
            }
        }

        $Cache[$DirectoryKey] = $resolved
        return $resolved
    }

    function Get-MsiAddInInfo {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$RegistryRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$ProgIdRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$ClassRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$FileRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$ComponentRows
        )

        $comAddinPattern = '(?i)\\Office\\(?:(?<App>[A-Za-z0-9\.]+)\\)?Addins\\(?<ProgId>[^\\]+)$'
        $excelOpenPattern = '(?i)\\Office\\(?:[^\\]+\\)?Excel\\Options$'
        $excelOpenValuePattern = '(?i)^OPEN\d*$'

        $comAddinGroups = [ordered]@{}
        $excelOpenEntries = @()

        foreach ($row in $RegistryRows) {
            $key = $row.Key
            if (-NOT $key) { continue }

            if ($key -match $comAddinPattern) {
                $hostApp = if ($Matches['App']) { $Matches['App'] } else { "(all Office hosts)" }
                $progId = $Matches['ProgId']
                $groupKey = "$($row.Root)|$key"

                if (-NOT $comAddinGroups.Contains($groupKey)) {
                    $comAddinGroups[$groupKey] = [PSCustomObject]@{
                        RegistryRoot    = $RegistryRootMap["$($row.Root)"]
                        RegistryKey     = $key
                        HostApplication = $hostApp
                        ProgId          = $progId
                        FriendlyName    = $null
                        Description     = $null
                        LoadBehavior    = $null
                        Component       = $row.Component_
                        Clsid           = $null
                        BinaryFile      = $null
                        AddInType       = "COM add-in"
                    }
                }

                $entry = $comAddinGroups[$groupKey]
                switch ($row.Name) {
                    "FriendlyName" { $entry.FriendlyName = $row.Value }
                    "Description" { $entry.Description = $row.Value }
                    "LoadBehavior" { $entry.LoadBehavior = $row.Value }
                }
            }

            if (($key -match $excelOpenPattern) -and ($row.Name -match $excelOpenValuePattern)) {
                $excelOpenEntries += [PSCustomObject]@{
                    RegistryRoot = $RegistryRootMap["$($row.Root)"]
                    RegistryKey  = $key
                    ValueName    = $row.Name
                    Value        = $row.Value
                    Component    = $row.Component_
                    AddInType    = "Excel classic add-in (XLL/XLA/XLAM) loaded via the 'OPEN' key"
                }
            }
        }

        foreach ($entry in $comAddinGroups.Values) {
            $progIdRow = $ProgIdRows | Where-Object { $_.ProgId -eq $entry.ProgId } | Select-Object -First 1
            if (-NOT $progIdRow) { continue }

            $entry.Clsid = $progIdRow.Class_
            $classRow = $ClassRows | Where-Object { $_.CLSID -eq $progIdRow.Class_ } | Select-Object -First 1
            if (-NOT $classRow) { continue }

            if (-NOT $entry.Description) { $entry.Description = $classRow.Description }

            $componentRow = $ComponentRows | Where-Object { $_.Component -eq $classRow.Component_ } | Select-Object -First 1
            if (-NOT $componentRow) { continue }

            $fileRow = $FileRows | Where-Object { $_.Component_ -eq $componentRow.Component } | Select-Object -First 1
            if ($fileRow) {
                $nameParts = $fileRow.FileName -split '\|'
                $entry.BinaryFile = if ($nameParts.Count -ge 2) { $nameParts[1] } else { $nameParts[0] }
            }
        }

        return [PSCustomObject]@{
            HasComAddIns          = (@($comAddinGroups.Values).Count -gt 0)
            HasExcelOpenKeyAddIns = (@($excelOpenEntries).Count -gt 0)
            ComAddIns             = @($comAddinGroups.Values)
            ExcelOpenKeyAddIns    = @($excelOpenEntries)
        }
    }

    function Get-MsiHkcuRegistryHives {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$RegistryRows,
            [Parameter(Mandatory = $false)][string]$AllUsersPropertyValue
        )

        # Root: 0=HKCR, 1=HKCU, 2=HKLM, 3=HKU, -1=HKCU or HKLM depending on ALLUSERS at real install time.
        # ALLUSERS set (usually "1") means a per-machine install, so -1 resolves to HKLM; otherwise it's
        # a per-user install and -1 resolves to HKCU. This is the same rule Windows Installer itself uses,
        # but note ALLUSERS can still be overridden on the real msiexec command line, so treat -1 rows as
        # "likely" rather than certain.
        $isPerMachineInstall = -NOT [string]::IsNullOrEmpty($AllUsersPropertyValue)

        $hkcuGroups = [ordered]@{}

        foreach ($row in $RegistryRows) {
            $isHkcu = $false
            $rootNote = $null

            if ($row.Root -eq "1") {
                $isHkcu = $true
            }
            elseif ($row.Root -eq "-1" -and -NOT $isPerMachineInstall) {
                $isHkcu = $true
                $rootNote = "Root is -1 (HKCU/HKLM depending on ALLUSERS); ALLUSERS is not set in this MSI's Property table, so this resolves to HKCU by default, but can be overridden at install time."
            }

            if (-NOT $isHkcu) { continue }

            $key = $row.Key
            if (-NOT $key) { continue }

            if (-NOT $hkcuGroups.Contains($key)) {
                $hkcuGroups[$key] = [PSCustomObject]@{
                    RegistryHive = "HKEY_CURRENT_USER"
                    RegistryKey  = $key
                    ValueNames   = @()
                    Components   = @()
                    Note         = $rootNote
                }
            }

            $entry = $hkcuGroups[$key]
            $valueName = if ($row.Name) { $row.Name } else { "@ (default value)" }
            if ($entry.ValueNames -notcontains $valueName) {
                $entry.ValueNames += $valueName
            }
            if ($row.Component_ -and ($entry.Components -notcontains $row.Component_)) {
                $entry.Components += $row.Component_
            }
        }

        return @($hkcuGroups.Values)
    }

    # --- Custom action / Binary stream add-in detection ------------------------------------------------
    #
    # Registration performed by a Custom Action (a VBScript/JScript or compiled DLL/EXE stored as an opaque
    # stream in the Binary table) is invisible to the declarative Registry table above. Record.ReadStream()
    # does not work reliably through PowerShell's late-bound COM interop (it throws DISP_E_BADINDEX in
    # testing), so Binary streams are instead extracted with Database.Export(), which writes real files to
    # disk (verified byte-for-byte, e.g. a DLL stream reads back with a correct "MZ" header) - still entirely
    # read-only against the source MSI. The extracted files are string-scanned like the classic "strings"
    # utility, then deleted. Because the code itself is never executed, a match only proves the relevant
    # words/paths are embedded in the binary/script, not the final runtime value - these are reported as
    # "likely" hints, not confirmed values.

    function Get-MsiPrintableStrings {
        param(
            [Parameter(Mandatory = $true)][byte[]]$Bytes,
            # 4, not 5: the single most important literal for this whole feature is "OPEN" itself (4 chars) -
            # compiled code commonly stores it as a standalone constant, concatenated with the key path at runtime.
            [int]$MinLength = 4
        )

        $pattern = "[\x20-\x7E]{$MinLength,}"
        $found = New-Object System.Collections.Generic.HashSet[string]

        $ascii = [System.Text.Encoding]::ASCII.GetString($Bytes)
        [regex]::Matches($ascii, $pattern) | ForEach-Object { [void]$found.Add($_.Value) }

        # Compiled DLL/EXE custom actions commonly store registry paths as wide (UTF-16LE) string literals
        $unicode = [System.Text.Encoding]::Unicode.GetString($Bytes)
        [regex]::Matches($unicode, $pattern) | ForEach-Object { [void]$found.Add($_.Value) }

        return $found
    }

    function Find-MsiAddInHints {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Strings
        )

        $hints = @()
        foreach ($s in $Strings) {
            if ([string]::IsNullOrWhiteSpace($s)) { continue }

            # Compiled code very often stores the "Office\<version>\" prefix and the "\Excel\Options" suffix as
            # SEPARATE string literals, concatenated at runtime - so the "Office\...\" prefix must be optional here,
            # not required, or the path-fragment-only case (just "\Excel\Options") is silently missed.
            if ($s -match '(?i)(Office\\[^\\]*\\)?Excel\\Options\b') {
                $hints += [PSCustomObject]@{ Category = "ExcelOptionsRegistryPath"; Match = $s }
            }
            elseif ($s -match '(?i)(Office\\[^\\]*\\)?Excel\\Add-in Manager\b') {
                $hints += [PSCustomObject]@{ Category = "ExcelAddInManagerPath"; Match = $s }
            }
            elseif ($s -match '(?i)Office\\(?:[A-Za-z0-9\.]+\\)?Addins\\') {
                $hints += [PSCustomObject]@{ Category = "ComAddinRegistryPath"; Match = $s }
            }

            if ($s -match '^(?i)OPEN\d*$') {
                $hints += [PSCustomObject]@{ Category = "ExcelOpenValueName"; Match = $s }
            }

            if ($s -match '(?i)\.(xll|xla|xlam)("|\\|$)') {
                $hints += [PSCustomObject]@{ Category = "ExcelAddInFilePath"; Match = $s }
            }

            if ($s -match '(?i)^Reg(CreateKey|SetValue|OpenKey|DeleteValue|DeleteKey)') {
                $hints += [PSCustomObject]@{ Category = "RegistryApiUsage"; Match = $s }
            }

            if ($s -match '(?i)(WScript\.Shell|RegWrite|regedit|reg\.exe\s+add)') {
                $hints += [PSCustomObject]@{ Category = "ScriptRegistryUsage"; Match = $s }
            }

            if ($s -match '\[[A-Za-z_][A-Za-z0-9_]*\][^\[\]]*\.(?i:xll|xla|xlam)') {
                $hints += [PSCustomObject]@{ Category = "FormattedAddInPath"; Match = $s }
            }
        }

        return $hints
    }

    # Categories specific enough on their own to mean something (an Excel/Office registry path fragment, or a
    # reference to an actual .xll/.xla/.xlam file). "ExcelOpenValueName" ("OPEN"), "RegistryApiUsage", and
    # "ScriptRegistryUsage" are common words/API names found in huge numbers of unrelated binaries (dialog
    # text, fopen/RegOpenKeyEx, etc.) - on their own they're just noise. A binary/action is only worth
    # surfacing when it has at least one strong hint; weak hints are kept alongside a strong one as context.
    $StrongAddInHintCategories = @("ExcelOptionsRegistryPath", "ExcelAddInManagerPath", "ComAddinRegistryPath", "ExcelAddInFilePath", "FormattedAddInPath")

    function Get-ConfidentAddInHints {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Hints
        )

        $hasStrongHint = @($Hints | Where-Object { $StrongAddInHintCategories -contains $_.Category }).Count -gt 0
        if (-NOT $hasStrongHint) { return @() }

        return @($Hints)
    }

    function Get-MsiBinaryNames {
        param([Parameter(Mandatory = $true)]$Database)

        if (-NOT (Test-MsiTableExists -Database $Database -TableName "Binary")) {
            return @()
        }

        $view = $Database.OpenView("SELECT ``Name`` FROM ``Binary``")
        [void]$view.Execute()

        $names = @()
        while ($true) {
            $record = $view.Fetch()
            if (-NOT $record) { break }
            $names += $record.StringData(1)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record)
        }

        [void]$view.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view)
        return $names
    }

    function Get-MsiBinaryStreamFindings {
        param(
            [Parameter(Mandatory = $true)]$Database,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$BinaryNames,
            [Parameter(Mandatory = $true)][string]$ExportFolder,
            [Parameter(Mandatory = $true)][long]$MaxBytes
        )

        if ($BinaryNames.Count -eq 0) { return @() }

        try {
            New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
            [void]$Database.Export("Binary", $ExportFolder, "Binary.idt")
        }
        catch {
            Write-Warning "Could not export Binary table streams for scanning: $_"
            return @()
        }

        $binarySubfolder = Join-Path $ExportFolder "Binary"
        if (-NOT (Test-Path -Path $binarySubfolder)) { return @() }

        $results = @()
        foreach ($file in Get-ChildItem -Path $binarySubfolder -File) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            if ($file.Length -gt $MaxBytes) {
                $results += [PSCustomObject]@{
                    BinaryName = $name
                    SizeBytes  = $file.Length
                    Scanned    = $false
                    SkipReason = "Stream exceeds MaxBinaryStreamScanBytes ($MaxBytes bytes); skipped for performance. Re-run with a higher -MaxBinaryStreamScanBytes to include it."
                    Hints      = @()
                }
                continue
            }

            $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
            $strings = @(Get-MsiPrintableStrings -Bytes $bytes -MinLength 4)
            $hints = @(Get-ConfidentAddInHints -Hints @(Find-MsiAddInHints -Strings $strings))

            $results += [PSCustomObject]@{
                BinaryName = $name
                SizeBytes  = $file.Length
                Scanned    = $true
                SkipReason = $null
                Hints      = $hints
            }
        }

        return $results
    }

    function Get-MsiCustomActionFindings {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$CustomActionRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$BinaryStreamFindings
        )

        $binaryFindingsByName = @{}
        foreach ($bf in $BinaryStreamFindings) { $binaryFindingsByName[$bf.BinaryName] = $bf }

        $results = @()
        foreach ($ca in $CustomActionRows) {
            $type = 0
            [void][int]::TryParse($ca.Type, [ref]$type)

            $baseType = $type -band 0x7
            $isDeferred = ($type -band 0x400) -ne 0
            $isCommit = ($type -band 0x200) -ne 0
            $isRollback = ($type -band 0x100) -ne 0

            $typeLabel = switch ($baseType) {
                1 { "DLL" }
                2 { "EXE" }
                3 { "TextData/Property" }
                5 { "JScript" }
                6 { "VBScript" }
                default { "Type $baseType" }
            }
            $timing = if ($isRollback) { "Rollback" } elseif ($isCommit) { "Commit" } elseif ($isDeferred) { "Deferred" } else { "Immediate" }

            $targetHints = @(Get-ConfidentAddInHints -Hints @(Find-MsiAddInHints -Strings @($ca.Target, $ca.Source)))

            # Link by name match against the Binary table rather than gating strictly on the decoded
            # "sourceIsBinary" bit: the msidbCustomActionType bit layout is not fully certain from the SDK docs
            # alone, and a missed link (false negative) is worse than an occasional coincidental name match.
            # $binaryFinding.Hints is already confidence-filtered by Get-MsiBinaryStreamFindings.
            $binaryFinding = $null
            $binaryHints = @()
            if ($ca.Source -and $binaryFindingsByName.ContainsKey($ca.Source)) {
                $binaryFinding = $binaryFindingsByName[$ca.Source]
                $binaryHints = @($binaryFinding.Hints)
            }

            $allHints = @($targetHints + $binaryHints)
            if ($allHints.Count -eq 0) { continue }

            $results += [PSCustomObject]@{
                Action           = $ca.Action
                Type             = $ca.Type
                TypeLabel        = $typeLabel
                Timing           = $timing
                Source           = $ca.Source
                Target           = $ca.Target
                LinkedBinaryName = if ($binaryFinding) { $ca.Source } else { $null }
                BinaryScanned    = if ($binaryFinding) { $binaryFinding.Scanned } else { $null }
                BinarySkipReason = if ($binaryFinding) { $binaryFinding.SkipReason } else { $null }
                Hints            = $allHints
            }
        }

        return $results
    }

    function Get-MsiInstalledAddInFiles {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$FileInventory
        )

        return @($FileInventory | Where-Object { $_.FileName -match '(?i)\.(xll|xla|xlam)$' })
    }

    function Get-MsiDriverInfo {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$ServiceInstallRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$FileRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$OdbcDriverRows
        )

        $driverServices = @($ServiceInstallRows | ForEach-Object {
                $st = 0
                [void][int]::TryParse($_.ServiceType, [ref]$st)
                if (($st -band 0x3) -ne 0) {
                    [PSCustomObject]@{
                        Name        = $_.Name
                        DisplayName = $_.DisplayName
                        ServiceType = $_.ServiceType
                        DriverKind  = if (($st -band 0x1) -ne 0) { "Kernel driver" } else { "File system driver" }
                        StartType   = $_.StartType
                        Component   = $_.Component_
                        Description = $_.Description
                    }
                }
            } | Where-Object { $_ })

        $driverFiles = @($FileRows | ForEach-Object {
                $parts = $_.FileName -split '\|'
                if ($parts | Where-Object { $_ -match '(?i)\.sys$' }) {
                    $longName = if ($parts.Count -ge 2) { $parts[1] } else { $parts[0] }
                    [PSCustomObject]@{
                        FileName  = $longName
                        Component = $_.Component_
                    }
                }
            } | Where-Object { $_ })

        return [PSCustomObject]@{
            DriverServices = $driverServices
            DriverSysFiles = $driverFiles
            OdbcDrivers    = @($OdbcDriverRows)
        }
    }

    function Get-MsiFileInventory {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$FileRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$ComponentRows,
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][hashtable]$DirectoryLookup,
            [Parameter(Mandatory = $true)][hashtable]$ResolvedProperties
        )

        $componentByKey = @{}
        foreach ($c in $ComponentRows) { $componentByKey[$c.Component] = $c }

        $dirCache = New-Object 'System.Collections.Generic.Dictionary[string,string]'

        $inventory = foreach ($file in $FileRows) {
            $nameParts = $file.FileName -split '\|'
            $longName = if ($nameParts.Count -ge 2) { $nameParts[1] } else { $nameParts[0] }

            $resolvedDir = $null
            $component = $componentByKey[$file.Component_]
            if ($component) {
                $resolvedDir = Resolve-MsiDirectoryPath -DirectoryLookup $DirectoryLookup -DirectoryKey $component.Directory_ -Cache $dirCache -ResolvedProperties $ResolvedProperties
            }

            [PSCustomObject]@{
                FileKey              = $file.File
                FileName             = $longName
                Component            = $file.Component_
                SizeBytes            = $file.FileSize
                Version              = $file.Version
                Sequence             = $file.Sequence
                ApproximateDirectory = $resolvedDir
                ApproximateFullPath  = if ($resolvedDir) { "$resolvedDir$longName" } else { $null }
            }
        }

        return @($inventory)
    }

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
    }
    catch {
        Throw "Unable to create the WindowsInstaller.Installer COM object. Windows Installer may not be registered on this machine. $_"
    }

    $script:fileCounter = 0
}

process {
    foreach ($path in $MsiPath) {
        $script:fileCounter++

        $resolvedPath = $null
        try {
            $resolvedPath = (Resolve-Path -Path $path).ProviderPath
        }
        catch {
            Write-Warning "File not found: $path"
            continue
        }

        if (-NOT $Quiet) {
            Write-Output "Scanning $resolvedPath ..."
        }

        $db = $null
        try {
            $db = $installer.OpenDatabase($resolvedPath, $MSI_OPEN_READONLY)

            $allTables = @((Get-MsiTableRows -Database $db -TableName "_Tables") | ForEach-Object { $_.Name })

            $registryRows = @(Get-MsiTableRows -Database $db -TableName "Registry")
            $progIdRows = @(Get-MsiTableRows -Database $db -TableName "ProgId")
            $classRows = @(Get-MsiTableRows -Database $db -TableName "Class")
            $fileRows = @(Get-MsiTableRows -Database $db -TableName "File")
            $componentRows = @(Get-MsiTableRows -Database $db -TableName "Component")
            $directoryRows = @(Get-MsiTableRows -Database $db -TableName "Directory")
            $serviceInstallRows = @(Get-MsiTableRows -Database $db -TableName "ServiceInstall")
            $odbcDriverRows = @(Get-MsiTableRows -Database $db -TableName "ODBCDriver")
            $featureRows = @(Get-MsiTableRows -Database $db -TableName "Feature")
            $featureComponentRows = @(Get-MsiTableRows -Database $db -TableName "FeatureComponents")
            $customActionRows = @(Get-MsiTableRows -Database $db -TableName "CustomAction")
            $shortcutRows = @(Get-MsiTableRows -Database $db -TableName "Shortcut")
            $environmentRows = @(Get-MsiTableRows -Database $db -TableName "Environment")
            $iniFileRows = @(Get-MsiTableRows -Database $db -TableName "IniFile")
            $upgradeRows = @(Get-MsiTableRows -Database $db -TableName "Upgrade")
            $launchConditionRows = @(Get-MsiTableRows -Database $db -TableName "LaunchCondition")
            $mediaRows = @(Get-MsiTableRows -Database $db -TableName "Media")

            $directoryLookup = @{}
            foreach ($d in $directoryRows) {
                $directoryLookup[$d.Directory] = @{ Parent = $d.Directory_Parent; DefaultDir = $d.DefaultDir }
            }

            $propertiesTable = Get-MsiPropertiesTable -Database $db
            $propertyRows = @($propertiesTable.PSObject.Properties | ForEach-Object { [PSCustomObject]@{ Property = $_.Name; Value = $_.Value } })
            $resolvedProperties = Get-MsiResolvedProperties -PropertyRows $propertyRows -CustomActionRows $customActionRows

            $addInInfo = Get-MsiAddInInfo -RegistryRows $registryRows -ProgIdRows $progIdRows -ClassRows $classRows -FileRows $fileRows -ComponentRows $componentRows
            $driverInfo = Get-MsiDriverInfo -ServiceInstallRows $serviceInstallRows -FileRows $fileRows -OdbcDriverRows $odbcDriverRows
            $fileInventory = @(Get-MsiFileInventory -FileRows $fileRows -ComponentRows $componentRows -DirectoryLookup $directoryLookup -ResolvedProperties $resolvedProperties)
            $hkcuHives = @(Get-MsiHkcuRegistryHives -RegistryRows $registryRows -AllUsersPropertyValue $propertiesTable.ALLUSERS)

            $installedAddInFiles = @(Get-MsiInstalledAddInFiles -FileInventory $fileInventory)

            $binaryStreamFindings = @()
            if (-NOT $SkipBinaryStreamScan) {
                $binaryNames = @(Get-MsiBinaryNames -Database $db)
                if ($binaryNames.Count -gt 0) {
                    $exportFolder = Join-Path ([System.IO.Path]::GetTempPath()) "msi-discovery-$([guid]::NewGuid())"
                    try {
                        $binaryStreamFindings = @(Get-MsiBinaryStreamFindings -Database $db -BinaryNames $binaryNames -ExportFolder $exportFolder -MaxBytes $MaxBinaryStreamScanBytes)
                    }
                    finally {
                        if (Test-Path -Path $exportFolder) {
                            Remove-Item -Path $exportFolder -Recurse -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
            }
            $customActionFindings = @(Get-MsiCustomActionFindings -CustomActionRows $customActionRows -BinaryStreamFindings $binaryStreamFindings)

            $addInInfo | Add-Member -NotePropertyName InstalledAddInFiles -NotePropertyValue $installedAddInFiles -Force
            $addInInfo | Add-Member -NotePropertyName HasInstalledAddInFiles -NotePropertyValue ($installedAddInFiles.Count -gt 0) -Force
            $addInInfo | Add-Member -NotePropertyName CustomActionFindings -NotePropertyValue $customActionFindings -Force
            $addInInfo | Add-Member -NotePropertyName HasCustomActionAddInHints -NotePropertyValue ($customActionFindings.Count -gt 0) -Force
            $addInInfo | Add-Member -NotePropertyName BinaryStreamScanSkipped -NotePropertyValue ([bool]$SkipBinaryStreamScan) -Force
            $addInInfo | Add-Member -NotePropertyName BinaryStreamFindings -NotePropertyValue $binaryStreamFindings -Force

            $report = [PSCustomObject]@{
                MsiPath      = $resolvedPath
                SummaryInfo  = Get-MsiSummaryInfo -Database $db
                Properties   = $propertiesTable
                AddIns       = $addInInfo
                HkcuRegistry = [PSCustomObject]@{
                    HasHkcuRegistryEntries = ($hkcuHives.Count -gt 0)
                    HiveCount              = $hkcuHives.Count
                    Hives                  = $hkcuHives
                }
                Drivers      = $driverInfo
                Files        = $fileInventory
                Registry     = @($registryRows)
                Components   = @($componentRows)
                Directories  = @($directoryRows)
                Features     = @($featureRows)
                FeatureComponents = @($featureComponentRows)
                CustomActions = @($customActionRows)
                Shortcuts    = @($shortcutRows)
                Environment  = @($environmentRows)
                IniFileEdits = @($iniFileRows)
                RelatedUpgrades = @($upgradeRows)
                LaunchConditions = @($launchConditionRows)
                Media        = @($mediaRows)
                AllTablesPresent = $allTables
                Error        = $null
            }

            if (-NOT $Quiet) {
                Write-Output ""
                Write-Output "Product: $($report.Properties.ProductName) $($report.Properties.ProductVersion) ($($report.Properties.Manufacturer))"
                Write-Output ""
                $hasHintEvidence = $report.AddIns.HasCustomActionAddInHints -or $report.AddIns.HasInstalledAddInFiles
                $excelAnswer = if ($report.AddIns.HasExcelOpenKeyAddIns) {
                    "YES (confirmed - declarative Registry table entry found, see values below)"
                }
                elseif ($hasHintEvidence) {
                    "LIKELY (no declarative Registry table entry, but a custom action / binary stream / installed .xll-.xla-.xlam file was found - see CustomActionFindings/InstalledAddInFiles below; the value can't be read statically because it's set by code, not the Registry table)"
                }
                else {
                    "NO (no evidence in the Registry table, custom actions, binary streams, or installed files)"
                }
                Write-Output "Q: Does this MSI install any Excel 'OPEN'-key add-ins? $excelAnswer"
                foreach ($a in $report.AddIns.ExcelOpenKeyAddIns) {
                    Write-Output "  - [$($a.RegistryRoot)] $($a.RegistryKey)\$($a.ValueName) points to: $($a.Value)"
                }
                foreach ($f in $report.AddIns.InstalledAddInFiles) {
                    Write-Output "  - Installed add-in file: $($f.FileName) -> $($f.ApproximateFullPath)"
                }
                foreach ($ca in $report.AddIns.CustomActionFindings) {
                    Write-Output "  - CustomAction '$($ca.Action)' [$($ca.TypeLabel), $($ca.Timing)] Source=$($ca.Source) Target=$($ca.Target)"
                    foreach ($h in $ca.Hints) {
                        Write-Output "      HINT [$($h.Category)]: $($h.Match)"
                    }
                    if ($ca.LinkedBinaryName -and $ca.BinaryScanned -eq $false) {
                        Write-Output "      NOTE: linked binary '$($ca.LinkedBinaryName)' was not scanned: $($ca.BinarySkipReason)"
                    }
                }
                if ($SkipBinaryStreamScan) {
                    Write-Output "  (Binary stream scan was skipped via -SkipBinaryStreamScan)"
                }
                Write-Output ""
                Write-Output "Q: Does this MSI write any HKEY_CURRENT_USER registry hives? $(if ($report.HkcuRegistry.HasHkcuRegistryEntries) { 'YES' } else { 'NO' }) (hive count: $($report.HkcuRegistry.HiveCount))"
                foreach ($h in $report.HkcuRegistry.Hives) {
                    Write-Output "  - HKEY_CURRENT_USER\$($h.RegistryKey)  [values: $($h.ValueNames -join ', ')]"
                    if ($h.Note) { Write-Output "      NOTE: $($h.Note)" }
                }
                Write-Output ""
                Write-Output "COM add-ins found: $($report.AddIns.ComAddIns.Count)"
                foreach ($a in $report.AddIns.ComAddIns) {
                    Write-Output "  - [$($a.HostApplication)] $($a.ProgId) LoadBehavior=$($a.LoadBehavior) Binary=$($a.BinaryFile)"
                }
                Write-Output "Driver services found: $($report.Drivers.DriverServices.Count), .sys files found: $($report.Drivers.DriverSysFiles.Count), ODBC drivers found: $($report.Drivers.OdbcDrivers.Count)"
                Write-Output "Files: $($report.Files.Count), Registry rows: $($report.Registry.Count), Custom actions: $($report.CustomActions.Count)"
                if ($report.CustomActions.Count -gt 0) {
                    Write-Output "NOTE: Custom actions can perform arbitrary logic (including installing add-ins/drivers) that is not visible in the tables above; review CustomActions manually."
                }
                Write-Output ""
            }

            if ($OutputJsonPath) {
                $targetJsonPath = $OutputJsonPath
                if ($script:fileCounter -gt 1) {
                    $ext = [System.IO.Path]::GetExtension($OutputJsonPath)
                    $base = $OutputJsonPath.Substring(0, $OutputJsonPath.Length - $ext.Length)
                    $targetJsonPath = "$base`_$($script:fileCounter)$ext"
                }
                ($report | ConvertTo-Json -Depth 15) | Set-Content -Path $targetJsonPath -Encoding utf8
                if (-NOT $Quiet) {
                    Write-Output "Saved full report to $targetJsonPath"
                }
            }

            Write-Output $report
        }
        catch {
            Write-Warning "Failed to scan $resolvedPath : $_"
            Write-Output ([PSCustomObject]@{
                    MsiPath = $resolvedPath
                    Error   = "$_"
                })
        }
        finally {
            if ($db) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
            }
        }
    }
}

end {
    if ($installer) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
