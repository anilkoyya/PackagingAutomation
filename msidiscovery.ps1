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
      - Drivers: ServiceInstall rows whose ServiceType flags mark them as a kernel or file-system driver,
        any File table entries with a .sys extension, and the ODBCDriver table when present.
      - Full file, registry, component, and directory inventories, with a best-effort (approximate) resolved
        install path for each file.
      - Summary Information stream, Property table, Feature tree, Custom Actions, Shortcuts, Environment,
        IniFile, and Upgrade table entries, plus a full list of every table present in the database so nothing
        vendor-specific gets silently missed.

    Because this only opens the database in read-only mode, it never modifies the MSI, never requires
    administrator rights, and never executes any installer or custom action code.
.PARAMETER MsiPath
    Path to one or more .msi files to scan. Accepts pipeline input, so it can be combined with Get-ChildItem
    to scan every MSI under a folder.
.PARAMETER OutputJsonPath
    Optional path to save the full report as JSON. When scanning multiple files via the pipeline, "_<n>" is
    appended before the extension for each additional file.
.PARAMETER Quiet
    Suppresses the human-readable console summary; the full report object is still returned on the pipeline.

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
    [switch]$Quiet
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

    function Resolve-MsiDirectoryPath {
        param(
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][hashtable]$DirectoryLookup,
            [Parameter(Mandatory = $true)][string]$DirectoryKey,
            [Parameter(Mandatory = $true)][System.Collections.Generic.Dictionary[string, string]]$Cache
        )

        if ($Cache.ContainsKey($DirectoryKey)) { return $Cache[$DirectoryKey] }

        if ($script:SpecialFolders.ContainsKey($DirectoryKey)) {
            $resolved = $script:SpecialFolders[$DirectoryKey]
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
            $resolved = "$($script:SpecialFolders['TARGETDIR'])$longName\"
        }
        else {
            $parentResolved = Resolve-MsiDirectoryPath -DirectoryLookup $DirectoryLookup -DirectoryKey $entry.Parent -Cache $Cache
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
            [Parameter(Mandatory = $true)][AllowEmptyCollection()][hashtable]$DirectoryLookup
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
                $resolvedDir = Resolve-MsiDirectoryPath -DirectoryLookup $DirectoryLookup -DirectoryKey $component.Directory_ -Cache $dirCache
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

            $addInInfo = Get-MsiAddInInfo -RegistryRows $registryRows -ProgIdRows $progIdRows -ClassRows $classRows -FileRows $fileRows -ComponentRows $componentRows
            $driverInfo = Get-MsiDriverInfo -ServiceInstallRows $serviceInstallRows -FileRows $fileRows -OdbcDriverRows $odbcDriverRows
            $fileInventory = Get-MsiFileInventory -FileRows $fileRows -ComponentRows $componentRows -DirectoryLookup $directoryLookup
            $hkcuHives = @(Get-MsiHkcuRegistryHives -RegistryRows $registryRows -AllUsersPropertyValue $propertiesTable.ALLUSERS)

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
                Write-Output "Q: Does this MSI install any Excel 'OPEN'-key add-ins? $(if ($report.AddIns.HasExcelOpenKeyAddIns) { 'YES' } else { 'NO' })"
                foreach ($a in $report.AddIns.ExcelOpenKeyAddIns) {
                    Write-Output "  - [$($a.RegistryRoot)] $($a.RegistryKey)\$($a.ValueName) points to: $($a.Value)"
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
