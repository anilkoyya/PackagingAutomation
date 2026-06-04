$exe = "C:\Temp\installer.exe"

$system32 = "C:\Windows\System32"
$syswow64 = "C:\Windows\SysWOW64"

# -----------------------------
# Step 1: Capture baseline file state
# -----------------------------
function Get-SystemSnapshot {
    param([string]$path)

    Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue |
    Select-Object FullName, Length, LastWriteTime |
    ForEach-Object {
        [PSCustomObject]@{
            Path = $_.FullName
            Size = $_.Length
            LastWrite = $_.LastWriteTime
        }
    }
}

Write-Host "Capturing BEFORE state..."
$beforeSys32 = Get-SystemSnapshot $system32
$beforeSysWOW = Get-SystemSnapshot $syswow64

# -----------------------------
# Step 2: Install silently
# -----------------------------
Write-Host "Installing package..."
Start-Process $exe -ArgumentList "/install /quiet /norestart" -Wait

# -----------------------------
# Step 3: Capture AFTER state
# -----------------------------
Write-Host "Capturing AFTER state..."
$afterSys32 = Get-SystemSnapshot $system32
$afterSysWOW = Get-SystemSnapshot $syswow64

# -----------------------------
# Step 4: Compare
# -----------------------------
Write-Host "Comparing differences..."

$diffSys32 = Compare-Object $beforeSys32 $afterSys32 -Property Path, Size, LastWrite -PassThru |
Where-Object { $_.SideIndicator -eq "=>" }

$diffSysWOW = Compare-Object $beforeSysWOW $afterSysWOW -Property Path, Size, LastWrite -PassThru |
Where-Object { $_.SideIndicator -eq "=>" }

# -----------------------------
# Step 5: Output results
# -----------------------------
$result = [PSCustomObject]@{
    System32Changes = $diffSys32
    SysWOW64Changes = $diffSysWOW
}

$result
