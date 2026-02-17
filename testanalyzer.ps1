<#
.SYNOPSIS
    MSI Analyzer - Analyzes MSI files for services, drivers, add-ins, custom actions, EOL components.
.DESCRIPTION
    Uses Windows Installer COM API to reliably read all MSI database tables.
.PARAMETER MsiPath
    Path to the MSI file.
.PARAMETER OutputPath
    Output HTML report path. Defaults to <MsiName>_Analysis.html.
.PARAMETER ListTables
    List all tables and exit.
.PARAMETER DumpTable
    Dump a specific table and exit.
.PARAMETER DumpRegistry
    Show all Registry entries color-coded for debugging.
.EXAMPLE
    .\Analyze-MSI.ps1 -MsiPath "C:\app.msi"
    .\Analyze-MSI.ps1 -MsiPath "app.msi" -ListTables
    .\Analyze-MSI.ps1 -MsiPath "app.msi" -DumpTable "CustomAction"
    .\Analyze-MSI.ps1 -MsiPath "app.msi" -DumpRegistry
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)][string]$MsiPath,
    [string]$OutputPath,
    [switch]$ListTables,
    [string]$DumpTable,
    [switch]$DumpRegistry
)

Add-Type -AssemblyName System.Web
$script:EscapeHtml = { param($s) [System.Web.HttpUtility]::HtmlEncode([string]$s) }

#region ===== MSI Database Access =====

function Read-MsiTable {
    param([object]$DB, [string]$Table)
    $rows = [System.Collections.ArrayList]::new()
    try {
        $view = $DB.GetType().InvokeMember("OpenView",[System.Reflection.BindingFlags]::InvokeMethod,$null,$DB,@("SELECT * FROM ``$Table``"))
        $view.GetType().InvokeMember("Execute",[System.Reflection.BindingFlags]::InvokeMethod,$null,$view,$null)
        $ci = $view.GetType().InvokeMember("ColumnInfo",[System.Reflection.BindingFlags]::GetProperty,$null,$view,@(0))
        $cc = [int]$ci.GetType().InvokeMember("FieldCount",[System.Reflection.BindingFlags]::GetProperty,$null,$ci,$null)
        $cn = [string[]]::new($cc)
        for($i=0;$i -lt $cc;$i++){$cn[$i]=$ci.GetType().InvokeMember("StringData",[System.Reflection.BindingFlags]::GetProperty,$null,$ci,@($i+1))}
        while($true){
            $rec = $view.GetType().InvokeMember("Fetch",[System.Reflection.BindingFlags]::InvokeMethod,$null,$view,$null)
            if($null -eq $rec){break}
            $row = [ordered]@{}
            for($i=0;$i -lt $cc;$i++){
                try{$row[$cn[$i]]=$rec.GetType().InvokeMember("StringData",[System.Reflection.BindingFlags]::GetProperty,$null,$rec,@($i+1))}
                catch{$row[$cn[$i]]=""}
            }
            [void]$rows.Add([PSCustomObject]$row)
        }
        $view.GetType().InvokeMember("Close",[System.Reflection.BindingFlags]::InvokeMethod,$null,$view,$null)
    } catch { Write-Verbose "Read-MsiTable $Table failed: $_" }
    return $rows
}

function Get-AllTableNames {
    param([object]$DB)
    $t = Read-MsiTable -DB $DB -Table "_Tables"
    return ($t | ForEach-Object { $_.Name } | Sort-Object)
}

#endregion

#region ===== Open Database =====
$MsiPath = $MsiPath.Trim('"',"'")
if(-not(Test-Path $MsiPath)){Write-Error "Not found: $MsiPath";exit 1}
$MsiPath = (Resolve-Path $MsiPath).Path
$fileSize = (Get-Item $MsiPath).Length
Write-Host "Analyzing: $MsiPath ($("{0:N0}" -f $fileSize) bytes)" -ForegroundColor Cyan

$installer = New-Object -ComObject WindowsInstaller.Installer
$DB = $installer.GetType().InvokeMember("OpenDatabase",[System.Reflection.BindingFlags]::InvokeMethod,$null,$installer,@($MsiPath,0))
#endregion

#region ===== Utility Modes =====
if($ListTables){
    $tables = Get-AllTableNames -DB $DB
    Write-Host "`nTables ($($tables.Count)):" -ForegroundColor Yellow
    foreach($t in $tables){
        $r = Read-MsiTable -DB $DB -Table $t
        Write-Host "  $t : $($r.Count) rows"
    }
    exit 0
}
if($DumpTable){
    $r = Read-MsiTable -DB $DB -Table $DumpTable
    Write-Host "`n$DumpTable : $($r.Count) rows" -ForegroundColor Yellow
    $r | Format-Table -AutoSize -Wrap
    exit 0
}
if($DumpRegistry){
    $rootMap = @{"-1"="HKCR";"0"="HKCR";"1"="HKCU";"2"="HKLM";"3"="HKU"}
    $rows = Read-MsiTable -DB $DB -Table "Registry"
    Write-Host "`nRegistry: $($rows.Count) entries" -ForegroundColor Yellow
    Write-Host ("-"*100)
    foreach($r in $rows){
        $root = if($rootMap.ContainsKey([string]$r.Root)){$rootMap[[string]$r.Root]}else{"?"}
        $key = [string]$r.Key; $kl = $key.ToLower()
        $color = "White"; $tag = ""
        if($kl -like '*addin*' -or $kl -like '*add-in*' -or $kl -like '*vsto*'){$color="Green";$tag=" [OFFICE ADD-IN]"}
        elseif($kl -like '*browser helper*'){$color="Yellow";$tag=" [BHO]"}
        elseif($kl -like '*clsid*'){$color="Cyan";$tag=" [COM]"}
        elseif($kl -like '*shellex*'){$color="Magenta";$tag=" [SHELL]"}
        Write-Host "  $root\$key$tag" -ForegroundColor $color
        Write-Host "    $($r.Name) = $($r.Value)" -ForegroundColor Gray
    }
    exit 0
}
#endregion

#region ===== Read All Tables =====
Write-Host "Reading tables..." -ForegroundColor Yellow
$tblProperty       = Read-MsiTable -DB $DB -Table "Property"
$tblServiceInstall = Read-MsiTable -DB $DB -Table "ServiceInstall"
$tblServiceControl = Read-MsiTable -DB $DB -Table "ServiceControl"
$tblRegistry       = Read-MsiTable -DB $DB -Table "Registry"
$tblFile           = Read-MsiTable -DB $DB -Table "File"
$tblCustomAction   = Read-MsiTable -DB $DB -Table "CustomAction"
$tblClass          = Read-MsiTable -DB $DB -Table "Class"
$tblTypeLib        = Read-MsiTable -DB $DB -Table "TypeLib"
$tblExtension      = Read-MsiTable -DB $DB -Table "Extension"
$tblEnvironment    = Read-MsiTable -DB $DB -Table "Environment"
$tblFont           = Read-MsiTable -DB $DB -Table "Font"
$tblLaunchCond     = Read-MsiTable -DB $DB -Table "LaunchCondition"
$allTableNames     = Get-AllTableNames -DB $DB
#endregion

#region ===== Build HTML Report =====
$E = $script:EscapeHtml
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Gather basic info
$prodName = ""; $prodVer = ""; $mfr = ""; $prodCode = ""; $upgradeCode = ""
foreach($p in $tblProperty){
    switch($p.Property){
        "ProductName"    {$prodName=$p.Value}
        "ProductVersion" {$prodVer=$p.Value}
        "Manufacturer"   {$mfr=$p.Value}
        "ProductCode"    {$prodCode=$p.Value}
        "UpgradeCode"    {$upgradeCode=$p.Value}
    }
}

# Severity badge helper
function Badge($sev){
    $colors = @{HIGH="#dc3545";MEDIUM="#fd7e14";LOW="#28a745";INFO="#6c757d";CRITICAL="#dc3545"}
    $c = if($colors.ContainsKey($sev)){$colors[$sev]}else{"#999"}
    return "<span style='display:inline-block;padding:2px 10px;border-radius:12px;font-size:.75rem;font-weight:600;color:#fff;background:$c'>$sev</span>"
}

$css = @"
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f8f9fa;color:#212529;line-height:1.6;padding:2rem;max-width:1200px;margin:auto}
.hd{background:linear-gradient(135deg,#1a237e,#1a73e8);color:#fff;padding:2rem;border-radius:12px;margin-bottom:2rem}
.hd h1{font-size:1.75rem;margin-bottom:.5rem}.hd .sb{opacity:.85;font-size:.95rem}
.gr{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1rem;margin-bottom:2rem}
.cd{background:#fff;border-radius:8px;padding:1.25rem;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.cd .lb{font-size:.78rem;color:#6c757d;text-transform:uppercase;letter-spacing:.05em}
.cd .vl{font-size:1.2rem;font-weight:700;margin-top:.25rem}
.cd .vl.hi{color:#dc3545}.cd .vl.ok{color:#28a745}
.sc{background:#fff;border-radius:8px;padding:1.5rem;margin-bottom:1.5rem;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.sc h2{font-size:1.15rem;margin-bottom:1rem;padding-bottom:.5rem;border-bottom:2px solid #dee2e6;display:flex;align-items:center;gap:.5rem}
.sc h3{font-size:1rem;margin:1rem 0 .5rem;color:#333}
table{width:100%;border-collapse:collapse;font-size:.85rem;margin-top:.5rem}
th{background:#f1f3f5;text-align:left;padding:8px 10px;font-weight:600;border-bottom:2px solid #dee2e6}
td{padding:8px 10px;border-bottom:1px solid #eee;vertical-align:top;word-break:break-word;max-width:400px}
tr:hover{background:#f8f9ff}
.pt td:first-child{font-weight:600;width:200px;color:#6c757d}
.nt{padding:.75rem 1rem;border-radius:4px;margin:.75rem 0;font-size:.9rem;border-left:4px solid}
.nt.dg{background:#f8d7da;border-color:#dc3545}.nt.wr{background:#fff3cd;border-color:#fd7e14}
.ft{text-align:center;color:#6c757d;font-size:.85rem;margin-top:2rem;padding-top:1rem;border-top:1px solid #dee2e6}
</style>
"@

$html = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>MSI Analysis - $(& $E $prodName)</title>$css</head><body>"

# Header
$html += "<div class='hd'><h1>MSI Analysis Report</h1>"
$html += "<div class='sb'><strong>$(& $E $prodName)</strong> v$(& $E $prodVer) &mdash; $(& $E $mfr)<br>"
$html += "Generated: $timestamp &bull; $(& $E (Split-Path $MsiPath -Leaf)) &bull; $("{0:N0}" -f $fileSize) bytes</div></div>"

# ===== Compute summary flags =====
$hasSvc  = ($tblServiceInstall.Count -gt 0 -or $tblServiceControl.Count -gt 0)
$hasDrv  = $false  # computed below
$hasCA   = ($tblCustomAction.Count -gt 0)

# Detect driver files
$driverFiles = [System.Collections.ArrayList]::new()
foreach($f in $tblFile){
    $fn = [string]$f.FileName; if($fn -match '\|'){$fn=($fn -split '\|')[-1]}
    $ext = [System.IO.Path]::GetExtension($fn).TrimStart('.').ToLower()
    if($ext -in @('sys','inf','cat')){
        [void]$driverFiles.Add(@{File=$fn;Type=$ext.ToUpper();Component=$f.Component_})
        $hasDrv = $true
    }
}

# Detect add-ins: group registry by key
$rootMap = @{"-1"="HKCR";"0"="HKCR";"1"="HKCU";"2"="HKLM";"3"="HKU"}
$officeAddins = [ordered]@{}  # key -> {App,Root,Component,Values=[list]}
$comRegs = [System.Collections.ArrayList]::new()
$shellExts = [System.Collections.ArrayList]::new()

foreach($r in $tblRegistry){
    $key = [string]$r.Key
    $kl = $key.ToLower()
    $rootStr = [string]$r.Root
    $rootName = if($rootMap.ContainsKey($rootStr)){$rootMap[$rootStr]}else{"Root($rootStr)"}
    $valName = [string]$r.Name
    $valData = [string]$r.Value
    $comp = [string]$r.Component_

    # Office add-in detection
    $isAddin = ($kl -like '*\addins\*') -or ($kl -like '*\addins') -or
               ($kl -like '*\add-ins\*') -or ($kl -like '*\add-ins') -or
               ($kl -like '*software\microsoft\vsto*') -or
               ($kl -like '*software\microsoft\office*addin*') -or
               ($kl -like '*software\wow6432node\microsoft\office*addin*')

    if($isAddin){
        if(-not $officeAddins.Contains($key)){
            $segs = $key -split '\\'
            $app = "Office"
            if($kl -like '*outlook*'){$app="Outlook"}
            elseif($kl -like '*excel*'){$app="Excel"}
            elseif($kl -like '*word*'){$app="Word"}
            elseif($kl -like '*powerpoint*'){$app="PowerPoint"}
            elseif($kl -like '*visio*'){$app="Visio"}
            elseif($kl -like '*access*'){$app="Access"}
            elseif($kl -like '*vsto*'){$app="VSTO"}
            $officeAddins[$key] = @{Name=$segs[-1];App=$app;Root=$rootName;Component=$comp;Vals=[System.Collections.ArrayList]::new()}
        }
        if($valName -or $valData){
            [void]$officeAddins[$key].Vals.Add("$valName = $valData")
        }
        continue
    }

    # COM
    if($kl -like '*clsid*' -and ($kl -like '*inprocserver*' -or $kl -like '*localserver*')){
        [void]$comRegs.Add(@{Root=$rootName;Key=$key;Name=$valName;Value=$valData;Component=$comp})
        continue
    }
    # Shell
    if($kl -like '*shellex*'){
        [void]$shellExts.Add(@{Root=$rootName;Key=$key;Name=$valName;Value=$valData;Component=$comp})
    }
}

# Secondary add-in detection: group by key, check for LoadBehavior
$regByKey = @{}
foreach($r in $tblRegistry){
    $key = [string]$r.Key
    if(-not $regByKey.ContainsKey($key)){$regByKey[$key]=[System.Collections.ArrayList]::new()}
    [void]$regByKey[$key].Add($r)
}
foreach($key in $regByKey.Keys){
    $kl = $key.ToLower()
    if($officeAddins.Contains($key)){continue}
    $entries = $regByKey[$key]
    $names = $entries | ForEach-Object {([string]$_.Name).ToLower()}
    if(($names -contains 'loadbehavior') -or (($names -contains 'friendlyname') -and ($names -contains 'manifest'))){
        $segs = $key -split '\\'
        $app = "Office"
        if($kl -like '*outlook*'){$app="Outlook"}
        elseif($kl -like '*excel*'){$app="Excel"}
        elseif($kl -like '*word*'){$app="Word"}
        elseif($kl -like '*powerpoint*'){$app="PowerPoint"}
        $first = $entries[0]
        $rootStr = [string]$first.Root
        $rootName = if($rootMap.ContainsKey($rootStr)){$rootMap[$rootStr]}else{"?"}
        $officeAddins[$key] = @{Name=$segs[-1];App=$app;Root=$rootName;Component=[string]$first.Component_;Vals=[System.Collections.ArrayList]::new()}
        foreach($e2 in $entries){
            $vn=[string]$e2.Name;$vd=[string]$e2.Value
            if($vn -or $vd){[void]$officeAddins[$key].Vals.Add("$vn = $vd")}
        }
    }
}

$hasAddin = ($officeAddins.Count -gt 0)
$hasCOM   = ($comRegs.Count -gt 0 -or $tblClass.Count -gt 0 -or $tblTypeLib.Count -gt 0)

# EOL Detection - file-based
$eolMap = @{}
function Add-EOL([string[]]$Files,[string]$Comp,[string]$Ver,[string]$EOL,[string]$Sev,[string]$Cat){
    foreach($f in $Files){$eolMap[$f.ToLower()]=@{Comp=$Comp;Ver=$Ver;EOL=$EOL;Sev=$Sev;Cat=$Cat}}
}

Add-EOL @('msvcr80.dll','msvcp80.dll','msvcm80.dll','atl80.dll','mfc80.dll','mfc80u.dll','msdia80.dll','vcomp80.dll') "Visual C++ 2005" "VC 8.0" "2016-04-12" "CRITICAL" "VC++ Runtime"
Add-EOL @('msvcr90.dll','msvcp90.dll','msvcm90.dll','atl90.dll','mfc90.dll','mfc90u.dll','msdia90.dll','vcomp90.dll') "Visual C++ 2008" "VC 9.0" "2018-04-10" "CRITICAL" "VC++ Runtime"
Add-EOL @('msvcr100.dll','msvcp100.dll','atl100.dll','mfc100.dll','mfc100u.dll','mfc100chs.dll','mfc100cht.dll','mfc100deu.dll','mfc100enu.dll','mfc100esn.dll','mfc100fra.dll','mfc100ita.dll','mfc100jpn.dll','mfc100kor.dll','mfc100rus.dll','msdia100.dll','vcomp100.dll','mfcm100.dll','mfcm100u.dll') "Visual C++ 2010" "VC 10.0" "2020-07-14" "HIGH" "VC++ Runtime"
Add-EOL @('msvcr110.dll','msvcp110.dll','atl110.dll','mfc110.dll','mfc110u.dll','mfc110chs.dll','mfc110cht.dll','mfc110deu.dll','mfc110enu.dll','mfc110esn.dll','mfc110fra.dll','mfc110ita.dll','mfc110jpn.dll','mfc110kor.dll','mfc110rus.dll','msdia110.dll','vcomp110.dll','mfcm110.dll','mfcm110u.dll') "Visual C++ 2012" "VC 11.0" "2023-06-13" "HIGH" "VC++ Runtime"
Add-EOL @('msvcr120.dll','msvcp120.dll','atl120.dll','mfc120.dll','mfc120u.dll','mfc120chs.dll','mfc120cht.dll','mfc120deu.dll','mfc120enu.dll','mfc120esn.dll','mfc120fra.dll','mfc120ita.dll','mfc120jpn.dll','mfc120kor.dll','mfc120rus.dll','msdia120.dll','vcomp120.dll','mfcm120.dll','mfcm120u.dll') "Visual C++ 2013" "VC 12.0" "2024-04-09" "MEDIUM" "VC++ Runtime"
Add-EOL @('java.exe','javaw.exe','javaws.exe','java.dll','jvm.dll','jawt.dll','jli.dll','jsound.dll','jaas_nt.dll','j2pcsc.dll','sunmscapi.dll','javac.exe','jar.exe','jps.exe','keytool.exe') "Oracle Java" "" "Licensing changed Jan 2023" "HIGH" "Java Runtime"
Add-EOL @('flash.ocx','flash32.ocx','flash64.ocx','pepflashplayer.dll') "Adobe Flash Player" "" "2020-12-31" "CRITICAL" "Deprecated Plugin"
Add-EOL @('agcore.dll','npctrl.dll','sllauncher.exe') "Microsoft Silverlight" "" "2021-10-12" "CRITICAL" "Deprecated Plugin"
Add-EOL @('msvbvm60.dll','vb6stkit.dll') "Visual Basic 6.0" "VB6" "2008-04-08" "HIGH" "Legacy Runtime"
Add-EOL @('libeay32.dll','ssleay32.dll') "OpenSSL 1.0.x" "1.0.x" "2019-12-31" "CRITICAL" "Deprecated Crypto"
Add-EOL @('sqlcese40.dll','sqlceqp40.dll','sqlce.dll') "SQL Server Compact" "" "2021-07-13" "HIGH" "Legacy Database"
Add-EOL @('dao360.dll','dao350.dll') "DAO" "" "Deprecated" "HIGH" "Legacy Database"

$eolRecs = @{HIGH="Recompile/upgrade to current supported version";"VC++ Runtime"="Recompile with VC++ 2015-2022 or bundle current VC++ Redistributable";"Java Runtime"="Migrate to Eclipse Temurin, Amazon Corretto, or Microsoft Build of OpenJDK";"Deprecated Plugin"="Remove entirely; migrate to HTML5/JavaScript";"Legacy Runtime"="Rewrite in .NET or modern language";"Deprecated Crypto"="Upgrade to OpenSSL 3.x";"Legacy Database"="Migrate to SQLite or SQL Server Express"}

$eolDetected = [ordered]@{}  # compName -> {info + files ArrayList}
foreach($f in $tblFile){
    $fn = [string]$f.FileName; if($fn -match '\|'){$fn=($fn -split '\|')[-1]}
    $fl = $fn.ToLower()
    $matched = $null

    if($eolMap.ContainsKey($fl)){$matched = $eolMap[$fl]}
    elseif($fl -like '*.jar'){$matched = @{Comp="Oracle Java";Ver="";EOL="Licensing changed Jan 2023";Sev="HIGH";Cat="Java Runtime"}}
    elseif($fl -like 'npswf*.dll'){$matched = @{Comp="Adobe Flash Player";Ver="";EOL="2020-12-31";Sev="CRITICAL";Cat="Deprecated Plugin"}}
    elseif($fl -like 'libssl-1_1*' -or $fl -like 'libcrypto-1_1*'){$matched = @{Comp="OpenSSL 1.1.x";Ver="1.1.x";EOL="2023-09-11";Sev="HIGH";Cat="Deprecated Crypto"}}
    elseif($fl -like 'msjet*.dll'){$matched = @{Comp="Jet Database";Ver="";EOL="Deprecated";Sev="MEDIUM";Cat="Legacy Database"}}

    if($null -ne $matched){
        $ck = $matched.Comp
        if(-not $eolDetected.Contains($ck)){
            $eolDetected[$ck] = @{Comp=$matched.Comp;Ver=$matched.Ver;EOL=$matched.EOL;Sev=$matched.Sev;Cat=$matched.Cat;Files=[System.Collections.ArrayList]::new()}
        }
        [void]$eolDetected[$ck].Files.Add("$fn ($($f.Component_))")
    }
}

$hasEOL = ($eolDetected.Count -gt 0)

# Custom action risk
$highRiskCA = 0
$exeTypes = @(1,2,17,18,34,50); $scriptTypes = @(5,6,21,22,37,38,53,54)
foreach($ca in $tblCustomAction){
    $base = [int]$ca.Type -band 0x3F
    if($base -in $exeTypes -or $base -in $scriptTypes){$highRiskCA++}
}

# Determine complexity
$complexity = "LOW"
$highAreas = [System.Collections.ArrayList]::new()
if($hasSvc){[void]$highAreas.Add("Services");$complexity="HIGH"}
if($hasDrv){[void]$highAreas.Add("Drivers");$complexity="HIGH"}
if($highRiskCA -gt 0){[void]$highAreas.Add("Custom Actions");$complexity="HIGH"}
if($hasEOL){[void]$highAreas.Add("EOL Components");if($complexity -ne "HIGH"){$complexity="HIGH"}}
if($complexity -ne "HIGH" -and ($hasAddin -or $hasCOM)){$complexity="MEDIUM"}

# ===== Summary Cards =====
$html += '<div class="gr">'
$cards = @(
    @("Services",$(if($hasSvc){"YES"}else{"No"})),
    @("Drivers",$(if($hasDrv){"YES"}else{"No"})),
    @("Add-ins",$(if($hasAddin){"YES ($($officeAddins.Count))"}else{"No"})),
    @("COM/ActiveX",$(if($hasCOM){"YES"}else{"No"})),
    @("EOL Components",$(if($hasEOL){"YES ($($eolDetected.Count))"}else{"No"})),
    @("Complexity",$complexity)
)
foreach($cd in $cards){
    $cls = if($cd[1] -match "YES|HIGH"){"hi"}else{"ok"}
    $ico = if($cls -eq "hi"){"&#9888;"}else{"&#10003;"}
    $html += "<div class='cd'><div class='lb'>$(& $E $cd[0])</div><div class='vl $cls'>$ico $(& $E $cd[1])</div></div>"
}
$html += '</div>'

# ===== Basic Info =====
$html += "<div class='sc'><h2>Basic Information</h2><table class='pt'>"
$infoRows = @(
    @("File", (Split-Path $MsiPath -Leaf)),
    @("File Size", "$("{0:N0}" -f $fileSize) bytes ($("{0:N2}" -f ($fileSize/1MB)) MB)"),
    @("ProductName",$prodName),@("ProductVersion",$prodVer),@("Manufacturer",$mfr),
    @("ProductCode",$prodCode),@("UpgradeCode",$upgradeCode),
    @("Tables",$allTableNames.Count)
)
foreach($r in $infoRows){$html += "<tr><td>$(& $E $r[0])</td><td>$(& $E $r[1])</td></tr>"}
$html += "</table></div>"

# ===== Services =====
if($hasSvc){
    $svcStartMap = @{"0"="Boot";"1"="System";"2"="Automatic";"3"="Manual";"4"="Disabled"}
    $svcTypeMap = @{"1"="Kernel Driver";"2"="FS Driver";"16"="Own Process";"32"="Share Process"}
    $html += "<div class='sc'><h2>Windows Services $(Badge 'HIGH')</h2>"
    $html += "<div class='nt dg'>This MSI installs/controls Windows services.</div>"

    if($tblServiceInstall.Count -gt 0){
        $html += "<h3>Service Installations ($($tblServiceInstall.Count))</h3>"
        foreach($s in $tblServiceInstall){
            $html += "<table class='pt'>"
            $html += "<tr><td>Service Name</td><td>$(& $E $s.Name)</td></tr>"
            $html += "<tr><td>Display Name</td><td>$(& $E $s.DisplayName)</td></tr>"
            $st = [string]$s.ServiceType; $html += "<tr><td>Service Type</td><td>$(if($svcTypeMap.ContainsKey($st)){$svcTypeMap[$st]}else{$st})</td></tr>"
            $sa = [string]$s.StartType; $html += "<tr><td>Start Type</td><td>$(if($svcStartMap.ContainsKey($sa)){$svcStartMap[$sa]}else{$sa})</td></tr>"
            $html += "<tr><td>Component</td><td>$(& $E $s.Component_)</td></tr>"
            $html += "<tr><td>Account</td><td>$(if($s.StartName){& $E $s.StartName}else{'LocalSystem'})</td></tr>"
            $html += "<tr><td>Description</td><td>$(& $E $s.Description)</td></tr>"
            $html += "</table><br>"
        }
    }
    if($tblServiceControl.Count -gt 0){
        $html += "<h3>Service Controls ($($tblServiceControl.Count))</h3><table>"
        $html += "<tr><th>Service</th><th>Event</th><th>Wait</th><th>Component</th></tr>"
        foreach($sc in $tblServiceControl){
            $ev=[int]$sc.Event;$evs=@()
            if($ev -band 1){$evs+="Start@Install"};if($ev -band 2){$evs+="Stop@Install"}
            if($ev -band 4){$evs+="Delete@Install"};if($ev -band 8){$evs+="Start@Uninstall"}
            if($ev -band 16){$evs+="Stop@Uninstall"};if($ev -band 32){$evs+="Delete@Uninstall"}
            $html += "<tr><td>$(& $E $sc.Name)</td><td>$($evs -join ', ')</td><td>$(if([int]$sc.Wait){'Yes'}else{'No'})</td><td>$(& $E $sc.Component_)</td></tr>"
        }
        $html += "</table>"
    }
    $html += "</div>"
}

# ===== Drivers =====
if($hasDrv){
    $html += "<div class='sc'><h2>Device Drivers $(Badge 'HIGH')</h2>"
    $html += "<div class='nt dg'>Contains device driver files.</div>"
    $html += "<table><tr><th>File</th><th>Type</th><th>Component</th></tr>"
    foreach($d in $driverFiles){$html += "<tr><td>$(& $E $d.File)</td><td>$($d.Type)</td><td>$(& $E $d.Component)</td></tr>"}
    $html += "</table></div>"
}

# ===== Add-ins & COM =====
if($hasAddin -or $hasCOM -or $shellExts.Count -gt 0 -or $tblTypeLib.Count -gt 0){
    $html += "<div class='sc'><h2>Add-ins &amp; COM Registration $(Badge 'MEDIUM')</h2>"

    # Office Add-ins
    if($officeAddins.Count -gt 0){
        $html += "<h3>Office Add-ins ($($officeAddins.Count))</h3>"
        $html += "<div class='nt wr'>Office add-ins integrate with Microsoft Office applications.</div>"
        $html += "<table><tr><th>Add-in Name</th><th>App</th><th>Root</th><th>Registry Key</th><th>Component</th><th>Values</th></tr>"
        foreach($key in $officeAddins.Keys){
            $a = $officeAddins[$key]
            $valStr = $a.Vals -join " | "
            $html += "<tr>"
            $html += "<td><strong>$(& $E $a.Name)</strong></td>"
            $html += "<td>$(& $E $a.App)</td>"
            $html += "<td>$(& $E $a.Root)</td>"
            $html += "<td>$(& $E $key)</td>"
            $html += "<td>$(& $E $a.Component)</td>"
            $html += "<td>$(& $E $valStr)</td>"
            $html += "</tr>"
        }
        $html += "</table>"
    }

    # COM from Class table
    if($tblClass.Count -gt 0){
        $html += "<h3>COM Classes ($($tblClass.Count))</h3><table>"
        $html += "<tr><th>CLSID</th><th>Context</th><th>Component</th><th>Description</th></tr>"
        foreach($c in ($tblClass | Select-Object -First 25)){
            $html += "<tr><td>$(& $E $c.CLSID)</td><td>$(& $E $c.Context)</td><td>$(& $E $c.Component_)</td><td>$(& $E $c.Description)</td></tr>"
        }
        $html += "</table>"
    }

    # COM from Registry
    if($comRegs.Count -gt 0){
        $html += "<h3>COM via Registry ($($comRegs.Count))</h3><table>"
        $html += "<tr><th>Root</th><th>Key</th><th>Name</th><th>Value</th></tr>"
        foreach($c in ($comRegs | Select-Object -First 20)){
            $html += "<tr><td>$(& $E $c.Root)</td><td>$(& $E $c.Key)</td><td>$(& $E $c.Name)</td><td>$(& $E $c.Value)</td></tr>"
        }
        $html += "</table>"
    }

    # Shell Extensions
    if($shellExts.Count -gt 0){
        $html += "<h3>Shell Extensions ($($shellExts.Count))</h3><table>"
        $html += "<tr><th>Root</th><th>Key</th><th>Name</th><th>Value</th></tr>"
        foreach($s in ($shellExts | Select-Object -First 15)){
            $html += "<tr><td>$(& $E $s.Root)</td><td>$(& $E $s.Key)</td><td>$(& $E $s.Name)</td><td>$(& $E $s.Value)</td></tr>"
        }
        $html += "</table>"
    }

    # TypeLibs
    if($tblTypeLib.Count -gt 0){
        $html += "<h3>Type Libraries ($($tblTypeLib.Count))</h3><table>"
        $html += "<tr><th>LibID</th><th>Version</th><th>Description</th><th>Component</th></tr>"
        foreach($t2 in $tblTypeLib){
            $html += "<tr><td>$(& $E $t2.LibID)</td><td>$(& $E $t2.Version)</td><td>$(& $E $t2.Description)</td><td>$(& $E $t2.Component_)</td></tr>"
        }
        $html += "</table>"
    }

    # File Extensions
    if($tblExtension.Count -gt 0){
        $html += "<h3>File Extensions ($($tblExtension.Count))</h3><table>"
        $html += "<tr><th>Extension</th><th>Component</th><th>ProgId</th></tr>"
        foreach($x in $tblExtension){
            $html += "<tr><td>$(& $E $x.Extension)</td><td>$(& $E $x.Component_)</td><td>$(& $E $x.ProgId_)</td></tr>"
        }
        $html += "</table>"
    }

    $html += "</div>"
}

# ===== Custom Actions =====
if($hasCA){
    $caDesc = @{1="DLL(Binary)";2="EXE(Binary)";5="JScript(Binary)";6="VBScript(Binary)";17="DLL(File)";18="EXE(File)";19="Error";21="JScript(File)";22="VBScript(File)";34="EXE(cmdline)";37="JScript(inline)";38="VBScript(inline)";50="EXE(workdir)";51="SetProperty";53="JScript(defer)";54="VBScript(defer)"}
    $sev = if($highRiskCA -gt 0){"HIGH"}else{"LOW"}
    $html += "<div class='sc'><h2>Custom Actions ($($tblCustomAction.Count)) $(Badge $sev)</h2>"
    if($highRiskCA -gt 0){$html += "<div class='nt dg'>$highRiskCA high-risk custom actions execute binary or script code.</div>"}
    $html += "<table><tr><th>Action</th><th>Type</th><th>Category</th><th>Risk</th><th>Source</th><th>Target</th></tr>"
    foreach($ca in $tblCustomAction){
        $tc=[int]$ca.Type;$base=$tc -band 0x3F
        $cat = if($caDesc.ContainsKey($base)){$caDesc[$base]}else{"Type $base"}
        $risk = if($base -in $exeTypes){"HIGH"}elseif($base -in $scriptTypes){"HIGH"}elseif($tc -band 0x800){"ELEVATED"}else{"LOW"}
        $rs = if($risk -match "HIGH"){"style='color:#dc3545;font-weight:700'"}else{""}
        $tgt = [string]$ca.Target; if($tgt.Length -gt 200){$tgt=$tgt.Substring(0,200)+"..."}
        $html += "<tr><td>$(& $E $ca.Action)</td><td>$tc</td><td>$cat</td><td $rs>$risk</td><td>$(& $E $ca.Source)</td><td>$(& $E $tgt)</td></tr>"
    }
    $html += "</table></div>"
}

# ===== EOL Components =====
if($hasEOL){
    $html += "<div class='sc'><h2>End-of-Life / Legacy Components ($($eolDetected.Count)) $(Badge 'HIGH')</h2>"
    $html += "<div class='nt dg'>Contains components that have reached End-of-Life. These no longer receive security updates.</div>"
    $html += "<table><tr><th>Component</th><th>Category</th><th>Severity</th><th>EOL Date</th><th>Files Found</th><th>Recommendation</th></tr>"
    foreach($ck in $eolDetected.Keys){
        $item = $eolDetected[$ck]
        $sevStyle = ""
        if($item.Sev -eq "CRITICAL"){$sevStyle="style='color:#dc3545;font-weight:700'"}
        elseif($item.Sev -eq "HIGH"){$sevStyle="style='color:#fd7e14;font-weight:700'"}
        $fileStr = "$($item.Files.Count) file(s): $(($item.Files | Select-Object -First 8) -join '; ')"
        $rec = if($eolRecs.ContainsKey($item.Cat)){$eolRecs[$item.Cat]}else{"Upgrade to current supported version"}
        $html += "<tr>"
        $html += "<td><strong>$(& $E $item.Comp)</strong></td>"
        $html += "<td>$(& $E $item.Cat)</td>"
        $html += "<td $sevStyle>$(& $E $item.Sev)</td>"
        $html += "<td>$(& $E $item.EOL)</td>"
        $html += "<td>$(& $E $fileStr)</td>"
        $html += "<td>$(& $E $rec)</td>"
        $html += "</tr>"
    }
    $html += "</table></div>"
}

# ===== Environment Variables =====
if($tblEnvironment.Count -gt 0){
    $html += "<div class='sc'><h2>Environment Variables $(Badge 'LOW')</h2><table>"
    $html += "<tr><th>Name</th><th>Value</th><th>Component</th></tr>"
    foreach($ev in $tblEnvironment){$html += "<tr><td>$(& $E $ev.Name)</td><td>$(& $E $ev.Value)</td><td>$(& $E $ev.Component_)</td></tr>"}
    $html += "</table></div>"
}

# ===== Launch Conditions =====
if($tblLaunchCond.Count -gt 0){
    $html += "<div class='sc'><h2>Launch Conditions $(Badge 'INFO')</h2><table>"
    $html += "<tr><th>Condition</th><th>Description</th></tr>"
    foreach($lc in $tblLaunchCond){$html += "<tr><td>$(& $E $lc.Condition)</td><td>$(& $E $lc.Description)</td></tr>"}
    $html += "</table></div>"
}

# ===== Files Summary =====
if($tblFile.Count -gt 0){
    $extCounts = @{}; $totalSz = [long]0
    foreach($f in $tblFile){
        $fn=[string]$f.FileName;if($fn -match '\|'){$fn=($fn -split '\|')[-1]}
        $ext=[System.IO.Path]::GetExtension($fn).TrimStart('.').ToLower()
        if(-not $ext){$ext="(none)"}
        $extCounts[$ext]=($extCounts[$ext] ?? 0)+1
        try{$totalSz+=[long]$f.FileSize}catch{}
    }
    $sorted = $extCounts.GetEnumerator()|Sort-Object Value -Descending|Select-Object -First 20
    $mx = ($sorted|Measure-Object -Property Value -Maximum).Maximum; if($mx -lt 1){$mx=1}
    $html += "<div class='sc'><h2>Files Summary $(Badge 'INFO')</h2>"
    $html += "<p><strong>$($tblFile.Count)</strong> files, <strong>$("{0:N2}" -f ($totalSz/1MB)) MB</strong></p>"
    $html += "<table><tr><th>Extension</th><th>Count</th><th></th></tr>"
    foreach($x in $sorted){
        $w=[Math]::Max(2,[int]($x.Value/$mx*200))
        $html += "<tr><td>.$($x.Key)</td><td>$($x.Value)</td><td><span style='display:inline-block;height:14px;background:#1a73e8;border-radius:3px;width:${w}px'></span></td></tr>"
    }
    $html += "</table></div>"
}

# ===== Fonts =====
if($tblFont.Count -gt 0){$html += "<div class='sc'><h2>Fonts $(Badge 'INFO')</h2><p>$($tblFont.Count) font(s)</p></div>"}

# ===== All Tables =====
$html += "<div class='sc'><h2>All Tables ($($allTableNames.Count))</h2><table><tr><th>Table</th><th>Rows</th></tr>"
foreach($t in $allTableNames){
    $r = Read-MsiTable -DB $DB -Table $t
    $html += "<tr><td>$(& $E $t)</td><td>$($r.Count)</td></tr>"
}
$html += "</table></div>"

# Footer
$html += "<div class='ft'>MSI Analyzer (PowerShell) &bull; $timestamp</div></body></html>"
#endregion

#region ===== Write Output =====
if(-not $OutputPath){
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($MsiPath)
    $OutputPath = Join-Path (Get-Location) "${baseName}_Analysis.html"
}
$html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
Write-Host "`nReport: $OutputPath" -ForegroundColor Green

Write-Host "`n=== SUMMARY ===" -ForegroundColor Yellow
$summaryLines = @(
    @("Contains Services",$(if($hasSvc){"YES"}else{"No"})),
    @("Contains Drivers",$(if($hasDrv){"YES"}else{"No"})),
    @("Office Add-ins",$officeAddins.Count),
    @("COM Registrations","$($comRegs.Count + $tblClass.Count)"),
    @("Custom Actions","$($tblCustomAction.Count) (High-Risk: $highRiskCA)"),
    @("EOL Components",$eolDetected.Count),
    @("Complexity",$complexity)
)
foreach($sl in $summaryLines){
    $v = [string]$sl[1]
    $color = if($v -match "YES|HIGH" -or ([int]::TryParse($v,[ref]$null) -and [int]$v -gt 0 -and $sl[0] -match "EOL")){"Red"}else{"White"}
    $marker = if($color -eq "Red"){"  !!!"}else{"     "}
    Write-Host "$marker $($sl[0]): $v" -ForegroundColor $color
}
if($highAreas.Count -gt 0){Write-Host "  !!! High-Severity: $($highAreas -join ', ')" -ForegroundColor Red}
#endregion