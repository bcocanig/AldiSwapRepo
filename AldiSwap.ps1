#Requires -Version 5.1
<#
.SYNOPSIS
    ALDI Laptop Swap tool - full rewrite (v21).

.DESCRIPTION
    Backs up a departing user's profile bits (OneNote notebooks + registry, Outlook
    profile registry, email signatures, Quick Access jump lists, Downloads, wallpaper,
    computer info, installed-app inventory) to OneDrive or an F: network share, then
    restores them onto a replacement laptop.

    This is a ground-up rewrite of v20reimport.ps1. Same job, rebuilt for stability:
      * One config block (no more 41 hardcoded paths).
      * Lazy OneDrive/F: resolution (no stale-at-load values).
      * Reusable helpers (reg export/import, robocopy, waits) - no copy-paste.
      * One data-driven backup + one restore pipeline (replaces 5 near-identical ones).
      * Every step wrapped: a failure logs and continues, never kills the session.
      * Approved verbs, UTF-8 everywhere, exit-code checks, elevation check.

.NOTES
    Original authors : Brandon Cocanig (11/23/2023), Chris Zeyen (Outlook fix).
    Rewrite          : v21 - 2026.
    Requirement      : runs on stock Windows PowerShell 5.1, no extra modules.

.PARAMETER Import
    Loads the functions only and skips the interactive menu. Intended for testing /
    dot-sourcing:  . .\AldiSwap.ps1 -Import
#>
param(
    [switch]$Import
)

# =====================================================================================
#region CONFIG  -- single source of truth (edit values here, nowhere else)
# =====================================================================================

$Script:OfficeVer = '16.0'   # Office major version used in registry paths

$Script:Config = [ordered]@{
    Root          = 'C:\Temp\LaptopTransferBackups'                       # local working folder
    FDriveRoot    = 'F:\usrnew\For IT Support\Laptop Swap Script\Backups' # divisional network backup
    SharePointUrl = 'https://asgportal-my.sharepoint.com/my'             # offline-restore download source
    # OneDrive folders are "OneDrive - <tenant>"; tried in this order (most specific first):
    OneDriveTenants = @('OneDrive - ALDI DX', 'OneDrive - ALDI-HOFER')
    Processes     = @{ OneNote = 'onenote'; OneNoteSender = 'ONENOTEM'; Outlook = 'OUTLOOK'; Explorer = 'explorer' }
}

# Derived sub-paths (built once from Root so nothing else hardcodes them)
$Script:Paths = [ordered]@{
    OneNoteJson        = Join-Path $Script:Config.Root 'OneNoteBooks.json'
    OneNoteCompareJson = Join-Path $Script:Config.Root 'OneNoteBooksCompare.json'
    OneNoteShortcuts   = Join-Path $Script:Config.Root 'OneNoteBooks_shortcuts'
    OneNoteReg         = Join-Path $Script:Config.Root 'OneNoteReg\OneNoteNotebooks.reg'
    OutlookRegDir      = Join-Path $Script:Config.Root 'OutlookReg'
    Signatures         = Join-Path $Script:Config.Root 'EmailSignatures'
    QuickAccess        = Join-Path $Script:Config.Root 'QuickAccessBackup'
    Downloads          = Join-Path $Script:Config.Root 'DownloadFiles'
    Wallpaper          = Join-Path $Script:Config.Root 'Wallpaper'
    Logs               = Join-Path $Script:Config.Root 'Logs'
    Trees              = Join-Path $Script:Config.Root 'Trees'
    ComputerInfo       = Join-Path $Script:Config.Root 'ComputerInfo.json'
    AppList            = Join-Path $Script:Config.Root 'InstalledApps.json'
}

# Registry keys (reg.exe form for export/import; PS-provider form for Test-Path)
$Script:Reg = @{
    OneNoteOpenExe = "HKEY_CURRENT_USER\Software\Microsoft\Office\$Script:OfficeVer\OneNote\OpenNotebooks"
    OutlookExe     = "HKEY_CURRENT_USER\Software\Microsoft\Office\$Script:OfficeVer\Outlook\Profiles"
    OutlookPS      = "HKCU:\Software\Microsoft\Office\$Script:OfficeVer\Outlook\Profiles"
    WallpaperPS    = 'HKCU:\Control Panel\Desktop'
}

$Script:OneNoteSchema = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

# Installed-app inventory map (Application -> install path / wildcard to probe)
$Script:AppMap = [ordered]@{
    'RedPrairie (MCH)'     = 'C:\Program Files (x86)\RedPrairie\MOCA\client'
    'Tableau Prep'         = 'C:\Program Files\Tableau Prep Builder'
    'Tableau Desktop'      = 'C:\Program Files\Tableau'
    'Spaceman'             = 'C:\Program Files\Spaceman'
    'Kofax'                = 'C:\Program Files (x86)\Kofax\AcrobatConnector'
    'Alteryx'              = 'C:\Program Files\Alteryx'
    'Git'                  = 'C:\Program Files\Git'
    'SSMS'                 = 'C:\Program Files (x86)\Microsoft SQL Server Management Studio*'
    'Anaconda'             = 'C:\Program Files\Anaconda3\python.exe'
    'Python'               = 'C:\Program Files\Python*\python.exe'
    '7-Zip'                = 'C:\Program Files\7-Zip\7z.exe'
    'Notepad++'            = 'C:\Program Files\Notepad++\notepad++.exe'
    'Visio'                = 'C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE'
    'Adobe Creative Cloud' = 'C:\Program Files (x86)\Adobe\Adobe Creative Cloud\CoreSync\CoreSync.exe'
    'Think-Cell'           = 'C:\Program Files\think-cell'
    'Visual Studio Code'   = 'C:\Program Files\Microsoft VS Code\Code.exe'
    'Visual Studio'        = 'C:\Program Files*\Microsoft Visual Studio\*\*\Common7\IDE\devenv.exe'
    'KeePass'              = 'C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe'
    'Kerberos'             = 'C:\Program Files (x86)\Kerberos\Kerberos.exe'
    'JDA Enterprise Client'= 'C:\Program Files (x86)\JDA\Enterprise*'
}

# Runtime state
$Script:Unattended  = $false   # when true, prompts auto-accept their default
$Script:LogFile     = $null
$Script:TranscriptOn= $false
$Script:OneNoteApp  = $null

# Sanity-check / progress state (populated per pipeline run)
$Script:Steps       = New-Object System.Collections.Generic.List[object]
$Script:StepIndex   = 0
$Script:StepTotal   = 0
$Script:RunStart    = $null

# Checklist marks (built from code points so the file stays pure-ASCII; avoids the
# PS 5.1 "no BOM -> mojibake" trap that bit the old script's em dashes).
$Script:Marks = @{
    Pass = [char]0x2713    # check
    Fail = [char]0x2717    # ballot X
    Skip = [char]0x2013    # en dash
    Warn = [char]0x0021    # !
}
$Script:MarkColor = @{ Pass = 'Green'; Fail = 'Red'; Skip = 'DarkYellow'; Warn = 'Yellow' }

#endregion

# =====================================================================================
#region LOGGING
# =====================================================================================

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK','STEP')][string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    $color = switch ($Level) { 'ERROR' {'Red'} 'WARN' {'Yellow'} 'OK' {'Green'} 'STEP' {'Cyan'} default {'Gray'} }
    Write-Host $line -ForegroundColor $color
}

function Start-SwapLog {
    New-DirIfMissing $Script:Paths.Logs
    $Script:LogFile = Join-Path $Script:Paths.Logs ("Log_{0}.txt" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
    try { Stop-Transcript | Out-Null } catch { }   # clear any orphaned transcript
    try {
        Start-Transcript -Path $Script:LogFile -Append | Out-Null
        $Script:TranscriptOn = $true
        Write-Log "Logging to $Script:LogFile" INFO
    } catch {
        $Script:TranscriptOn = $false
        Write-Log "Could not start transcript: $($_.Exception.Message)" WARN
    }
}

function Stop-SwapLog {
    if ($Script:TranscriptOn) {
        try { Stop-Transcript | Out-Null } catch { }
        $Script:TranscriptOn = $false
    }
}

#endregion

# =====================================================================================
#region CORE HELPERS
# =====================================================================================

function Set-ConsoleUtf8 {
    # Lets the Unicode checklist marks render. Guarded: harmless if the host refuses.
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
}

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-DirIfMissing {
    param([string]$Path)
    if ($Path -and -not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Confirm-Action {
    # Y/N prompt that honors -Unattended (returns the default without asking).
    param([string]$Message, [bool]$DefaultYes = $true)
    if ($Script:Unattended) { return $DefaultYes }
    $suffix = if ($DefaultYes) { '[Y/n]' } else { '[y/N]' }
    while ($true) {
        $r = (Read-Host "$Message $suffix").Trim().ToLower()
        if ($r -eq '')             { return $DefaultYes }
        if ($r -in @('y','yes'))   { return $true }
        if ($r -in @('n','no'))    { return $false }
        Write-Host 'Please answer y or n.'
    }
}

function Select-FromList {
    param([object[]]$Items, [scriptblock]$Display, [string]$Prompt = 'Select')
    $Items = @($Items)
    if ($Items.Count -eq 0) { return $null }
    if ($Script:Unattended -or $Items.Count -eq 1) { return $Items[0] }
    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host ("  {0}. {1}" -f $i, (& $Display $Items[$i]))
    }
    while ($true) {
        $r = (Read-Host $Prompt).Trim()
        if ($r -eq '') { return $Items[0] }
        if ($r -match '^\d+$' -and [int]$r -lt $Items.Count) { return $Items[[int]$r] }
        Write-Host "Enter a number 0-$($Items.Count - 1), or press Enter for the latest."
    }
}

function Wait-ForCondition {
    param(
        [Parameter(Mandatory)][scriptblock]$Condition,
        [int]$TimeoutSeconds = 120,
        [int]$PollSeconds = 2,
        [string]$Activity = 'Waiting'
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (& $Condition) { return $true }
        Start-Sleep -Seconds $PollSeconds
    }
    Write-Log "$Activity timed out after $TimeoutSeconds s." WARN
    return $false
}

function Stop-ProcessSafe {
    param([string]$Name, [int]$TimeoutSeconds = 30)
    if (-not (Get-Process -Name $Name -ErrorAction SilentlyContinue)) {
        Write-Log "$Name not running." INFO
        return $true
    }
    Write-Log "Stopping $Name..." INFO
    Get-Process -Name $Name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Wait-ForCondition -Condition { -not (Get-Process -Name $Name -ErrorAction SilentlyContinue) } `
                      -TimeoutSeconds $TimeoutSeconds -PollSeconds 1 -Activity "Closing $Name"
}

function Invoke-Robocopy {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [string[]]$Options = @('/E','/COPY:DAT','/R:1','/W:1','/NP','/NFL','/NDL')
    )
    if (-not (Test-Path -LiteralPath $Source)) { Write-Log "Robocopy source missing: $Source" WARN; return $false }
    New-DirIfMissing $Destination
    & robocopy $Source $Destination @Options | Out-Null
    $code = $LASTEXITCODE
    if ($code -ge 8) { Write-Log "Robocopy '$Source' -> '$Destination' FAILED (exit $code)." ERROR; return $false }
    Write-Log "Robocopy '$Source' -> '$Destination' ok (exit $code)." OK
    return $true
}

function Export-RegKey {
    param([Parameter(Mandatory)][string]$KeyPath, [Parameter(Mandatory)][string]$OutFile)
    New-DirIfMissing (Split-Path $OutFile -Parent)
    & reg.exe export $KeyPath $OutFile /y | Out-Null
    if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $OutFile)) {
        Write-Log "Exported $KeyPath -> $OutFile" OK; return $true
    }
    Write-Log "Failed to export $KeyPath (exit $LASTEXITCODE)." ERROR; return $false
}

function Import-RegKey {
    param([Parameter(Mandatory)][string]$RegFile)
    if (-not (Test-Path -LiteralPath $RegFile)) { Write-Log "Reg file not found: $RegFile" ERROR; return $false }
    & reg.exe import $RegFile | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Log "Imported $RegFile" OK; return $true }
    Write-Log "Failed to import $RegFile (exit $LASTEXITCODE)." ERROR; return $false
}

function Write-JsonFile {
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$Path, [int]$Depth = 5)
    New-DirIfMissing (Split-Path $Path -Parent)
    $Object | ConvertTo-Json -Depth $Depth | Set-Content -Path $Path -Encoding UTF8
}

function Read-JsonFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

#endregion

# =====================================================================================
#region LOCATION RESOLUTION  (lazy - resolved each time, never stale)
# =====================================================================================

function Resolve-OneDrive {
    # Find "OneDrive - <tenant>" under any profile matching this user (handles .ALDI-499).
    foreach ($tenant in $Script:Config.OneDriveTenants) {
        $profiles = Get-ChildItem -Path 'C:\Users' -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like "$env:USERNAME*" }
        foreach ($p in $profiles) {
            $candidate = Join-Path $p.FullName $tenant
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
    }
    return $null
}

function Resolve-FDrive {
    if (Test-Path -LiteralPath $Script:Config.FDriveRoot) { return $Script:Config.FDriveRoot }
    return $null
}

function Resolve-OutlookExe {
    $known = @(
        'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'
    )
    foreach ($p in $known) { if (Test-Path -LiteralPath $p) { return $p } }
    $appPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
    if (Test-Path $appPath) {
        $v = (Get-ItemProperty $appPath -ErrorAction SilentlyContinue).'(default)'
        if ($v -and (Test-Path -LiteralPath $v)) { return $v }
    }
    return $null
}

function Test-OneDriveReady {
    # Returns the OneDrive path, launching OneDrive and waiting for sign-in if needed.
    param([int]$TimeoutSeconds = 300)
    $od = Resolve-OneDrive
    if ($od) { Write-Log "OneDrive found: $od" OK; return $od }

    Write-Log 'OneDrive folder not found (not signed in?).' WARN
    $launcher = Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'
    if (Test-Path -LiteralPath $launcher) { Write-Log 'Launching OneDrive...' INFO; Start-Process $launcher }
    if (-not $Script:Unattended) { Write-Host 'Sign into OneDrive; it will be detected automatically.' }

    if (Wait-ForCondition -Condition { Resolve-OneDrive } -TimeoutSeconds $TimeoutSeconds -PollSeconds 3 -Activity 'OneDrive sign-in') {
        $od = Resolve-OneDrive
        Write-Log "OneDrive found: $od" OK
        return $od
    }
    return $null
}

#endregion

# =====================================================================================
#region HOUSEKEEPING
# =====================================================================================

function Clear-StaleWorkingFolder {
    $root = $Script:Config.Root
    if (-not (Test-Path -LiteralPath $root)) { Write-Log 'No stale working folder.' INFO; return }
    Write-Log "Existing working folder found: $root" WARN
    if (Confirm-Action 'Delete it before continuing?' $true) {
        try   { Remove-Item -Path $root -Recurse -Force -ErrorAction Stop; Write-Log "Deleted $root" OK }
        catch { Write-Log "Could not delete ${root}: $($_.Exception.Message)" ERROR }
    } else {
        Write-Log 'Keeping existing working folder.' INFO
    }
}

#endregion

# =====================================================================================
#region ONENOTE
# =====================================================================================

function Get-OneNoteApp {
    if ($Script:OneNoteApp) { return $Script:OneNoteApp }
    try {
        $interop = Get-ChildItem "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote" `
                   -Recurse -Filter '*.dll' -ErrorAction Stop | Select-Object -First 1
        if (-not $interop) { throw 'OneNote Interop assembly not found in the GAC.' }
        try { Add-Type -LiteralPath $interop.FullName -ErrorAction Stop } catch { }  # ignore "already loaded"
        $Script:OneNoteApp = New-Object Microsoft.Office.Interop.OneNote.ApplicationClass
        return $Script:OneNoteApp
    } catch {
        Write-Log "Unable to create OneNote COM object: $($_.Exception.Message)" ERROR
        return $null
    }
}

function Get-OneNoteNotebookList {
    $app = Get-OneNoteApp
    if (-not $app) { return @() }
    try {
        $xml = ''
        $app.GetHierarchy($null, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsNotebooks, [ref]$xml)
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($xml)
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('one', $Script:OneNoteSchema)
        $doc.SelectNodes('//one:Notebook', $ns) | ForEach-Object {
            [PSCustomObject]@{ Name = $_.GetAttribute('name'); Path = $_.GetAttribute('path') }
        }
    } catch {
        Write-Log "Failed to read OneNote hierarchy: $($_.Exception.Message)" ERROR
        return @()
    }
}

function Export-OneNoteNotebooks {
    param([string]$OutFile = $Script:Paths.OneNoteJson)
    $nb = @(Get-OneNoteNotebookList)
    Write-JsonFile -Object $nb -Path $OutFile
    Write-Log ("Recorded {0} notebook(s) -> {1}" -f $nb.Count, $OutFile) OK
    return $nb
}

function Export-OneNoteRegistry {
    Export-RegKey -KeyPath $Script:Reg.OneNoteOpenExe -OutFile $Script:Paths.OneNoteReg
}

function New-OneNoteShortcuts {
    $json = Read-JsonFile $Script:Paths.OneNoteJson
    if (-not $json) { Write-Log 'No notebook list to build shortcuts from.' WARN; return }
    $dir = $Script:Paths.OneNoteShortcuts
    New-DirIfMissing $dir
    $shell = New-Object -ComObject WScript.Shell
    $count = 0
    foreach ($item in $json) {
        $path = $item.Path
        if (-not $path) { continue }
        if ($path -like 'http*') {
            $sc = $shell.CreateShortcut((Join-Path $dir "$($item.Name).url"))
            $sc.TargetPath = "onenote:$path"; $sc.Save(); $count++
        } elseif (Test-Path -LiteralPath $path) {
            $toc = Get-ChildItem $path -Recurse -Filter 'Open Notebook.onetoc2' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($toc) {
                $folder = Split-Path (Split-Path $toc.FullName -Parent) -Leaf
                $sc = $shell.CreateShortcut((Join-Path $dir "$folder - Open Notebook.lnk"))
                $sc.TargetPath = $toc.FullName; $sc.Save(); $count++
            } else {
                Write-Log "No 'Open Notebook.onetoc2' under $path" WARN
            }
        }
    }
    Write-Log "Created $count OneNote shortcut(s) in $dir" OK
}

function Import-OneNoteRegistry {
    Stop-ProcessSafe $Script:Config.Processes.OneNote       | Out-Null
    Stop-ProcessSafe $Script:Config.Processes.OneNoteSender | Out-Null
    Import-RegKey -RegFile $Script:Paths.OneNoteReg
}

function Compare-OneNoteNotebook {
    $before = Read-JsonFile $Script:Paths.OneNoteJson
    if (-not $before) { Write-Log 'No pre-swap notebook list to compare against.' WARN; return }
    Write-Log 'Opening OneNote to re-check notebooks...' INFO
    Start-Process 'onenote.exe' -ErrorAction SilentlyContinue
    Wait-ForCondition -Condition { Get-Process -Name $Script:Config.Processes.OneNote -ErrorAction SilentlyContinue } `
                      -TimeoutSeconds 60 -PollSeconds 2 -Activity 'OneNote start' | Out-Null
    Start-Sleep -Seconds 5   # brief settle: OneNote enumerates notebooks asynchronously after launch
    $after   = @(Export-OneNoteNotebooks -OutFile $Script:Paths.OneNoteCompareJson)
    $missing = Compare-Object -ReferenceObject @($before.Name) -DifferenceObject @($after.Name) |
               Where-Object SideIndicator -eq '<=' | ForEach-Object InputObject
    if ($missing) {
        Write-Log 'Notebooks present before but MISSING after restore:' WARN
        foreach ($m in $missing) {
            $o = $before | Where-Object Name -eq $m
            Write-Log " - $($o.Name)  [$($o.Path)]" WARN
        }
        return 'WARN'
    }
    Write-Log 'All notebooks present.' OK
    return $true
}

#endregion

# =====================================================================================
#region OUTLOOK
# =====================================================================================

function Export-OutlookProfile {
    param([string]$Name = 'OldPcOutlook')
    Export-RegKey -KeyPath $Script:Reg.OutlookExe -OutFile (Join-Path $Script:Paths.OutlookRegDir "$Name.reg")
}

function Wait-ForOutlookProfile {
    param([int]$TimeoutSeconds = 300)
    Write-Log 'Waiting for Outlook to create its profile (open Outlook + sign in)...' INFO
    Wait-ForCondition -Condition { Test-Path $Script:Reg.OutlookPS } `
                      -TimeoutSeconds $TimeoutSeconds -PollSeconds 3 -Activity 'Outlook profile creation'
}

function Import-OutlookProfile {
    param([string]$Name = 'OldPcOutlook')
    Export-OutlookProfile -Name 'NewPcBackup' | Out-Null   # snapshot the new-PC profile first
    Stop-ProcessSafe $Script:Config.Processes.Outlook | Out-Null
    Import-RegKey -RegFile (Join-Path $Script:Paths.OutlookRegDir "$Name.reg")
}

function Repair-OutlookProfile {
    # Re-import the new-PC profile snapshot to undo a bad Outlook state after a restore.
    Stop-ProcessSafe $Script:Config.Processes.Outlook | Out-Null
    Import-RegKey -RegFile (Join-Path $Script:Paths.OutlookRegDir 'NewPcBackup.reg')
}

#endregion

# =====================================================================================
#region SIGNATURES / QUICK ACCESS / DOWNLOADS / WALLPAPER
# =====================================================================================

function Sync-Signature {
    param([Parameter(Mandatory)][ValidateSet('Backup','Restore')][string]$Direction)
    $live  = Join-Path $env:APPDATA 'Microsoft\Signatures'
    $store = $Script:Paths.Signatures
    if ($Direction -eq 'Backup') { $src = $live;  $dst = $store }
    else                         { $src = $store; $dst = $live  }
    Invoke-Robocopy -Source $src -Destination $dst
}

function Backup-QuickAccess {
    $src = Join-Path $env:APPDATA 'Microsoft\Windows\Recent\AutomaticDestinations'
    Invoke-Robocopy -Source $src -Destination $Script:Paths.QuickAccess
}

function Restore-QuickAccess {
    $dst = Join-Path $env:APPDATA 'Microsoft\Windows\Recent\AutomaticDestinations'
    $src = $Script:Paths.QuickAccess
    if (-not (Test-Path -LiteralPath $src)) { Write-Log 'No Quick Access backup to restore.' WARN; return $false }
    $ok = Invoke-Robocopy -Source $src -Destination $dst -Options @('/E','/COPY:DAT','/IS','/R:1','/W:1','/NP','/NFL','/NDL')
    Write-Log 'Restarting Explorer to refresh Quick Access...' INFO
    Stop-ProcessSafe $Script:Config.Processes.Explorer | Out-Null
    if (-not (Get-Process -Name $Script:Config.Processes.Explorer -ErrorAction SilentlyContinue)) {
        Start-Process explorer.exe
    }
    return $ok
}

function Backup-Downloads {
    param([double]$PromptThresholdGB = 1)
    $src = Join-Path $env:USERPROFILE 'Downloads'
    if (-not (Test-Path -LiteralPath $src)) { Write-Log 'No Downloads folder.' WARN; return $false }
    $bytes = (Get-ChildItem $src -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    $gb = [math]::Round((($bytes) / 1GB), 2)
    if ($gb -ge $PromptThresholdGB) {
        if (-not (Confirm-Action "Downloads is ${gb} GB. Back it up?" $true)) {
            Write-Log 'Skipped Downloads backup.' INFO; return $false
        }
    }
    New-DirIfMissing $Script:Paths.Downloads
    $zip = Join-Path $Script:Paths.Downloads ("downloadFiles_{0}.zip" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        Compress-Archive -Path (Join-Path $src '*') -DestinationPath $zip -CompressionLevel Fastest -Force -ErrorAction Stop
        Write-Log "Downloads zipped -> $zip" OK; return $true
    } catch {
        Write-Log "Downloads zip failed: $($_.Exception.Message)" ERROR; return $false
    }
}

function Backup-Wallpaper {
    # Backup only (per design): capture the current wallpaper file for reference.
    try { $wp = (Get-ItemProperty $Script:Reg.WallpaperPS -Name Wallpaper -ErrorAction Stop).Wallpaper }
    catch { $wp = $null }
    if ($wp -and (Test-Path -LiteralPath $wp)) {
        New-DirIfMissing $Script:Paths.Wallpaper
        Copy-Item -LiteralPath $wp -Destination $Script:Paths.Wallpaper -Force
        Write-Log "Wallpaper backed up: $wp" OK; return $true
    }
    Write-Log 'No wallpaper file found to back up.' WARN; return $false
}

#endregion

# =====================================================================================
#region SYSTEM INFO / INVENTORY
# =====================================================================================

function Get-DeviceSummary {
    # (Renamed from the old Get-ComputerInfo, which shadowed the built-in cmdlet.)
    $info = [ordered]@{
        'Device Name' = $env:COMPUTERNAME
        'User'        = $env:USERNAME
        'OS'          = (Get-CimInstance Win32_OperatingSystem).Caption
        'RAM (GB)'    = '{0:N2}' -f ((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB)
        'Storage'     = @(Get-CimInstance Win32_LogicalDisk | Where-Object DriveType -eq 3 |
                          ForEach-Object { "$($_.DeviceID) $([math]::Round($_.FreeSpace/1GB,2))GB free / $([math]::Round($_.Size/1GB,2))GB" })
        'Model'       = (Get-CimInstance Win32_ComputerSystem).Model
        'BIOS'        = (Get-CimInstance Win32_BIOS).SMBIOSBIOSVersion
        'Service Tag' = (Get-CimInstance Win32_BIOS).SerialNumber
    }
    $info.GetEnumerator() | ForEach-Object { Write-Log ("{0}: {1}" -f $_.Key, ($_.Value -join ', ')) INFO }
    Write-JsonFile -Object $info -Path $Script:Paths.ComputerInfo
    Write-Log "Device summary -> $($Script:Paths.ComputerInfo)" OK
}

function Get-InstalledAppList {
    $rows = foreach ($name in $Script:AppMap.Keys) {
        $hit = Get-Item -Path $Script:AppMap[$name] -ErrorAction SilentlyContinue | Select-Object -First 1
        $installed = if ($hit) { 'Yes' } else { 'No' }
        $path      = if ($hit) { $hit.FullName } else { '' }
        [PSCustomObject]@{ Application = $name; Installed = $installed; Path = $path }
    }
    $rows | Where-Object Installed -eq 'Yes' | Format-Table -AutoSize | Out-Host
    Write-JsonFile -Object $rows -Path $Script:Paths.AppList
    Write-Log "App inventory -> $($Script:Paths.AppList)" OK
}

function Save-FolderTree {
    param([Parameter(Mandatory)][string]$Directory, [Parameter(Mandatory)][string]$Label)
    if (-not (Test-Path -LiteralPath $Directory)) { Write-Log "Tree target missing: $Directory" WARN; return }
    New-DirIfMissing $Script:Paths.Trees
    $out = Join-Path $Script:Paths.Trees "$Label.txt"
    & tree.com $Directory /f | Out-File -FilePath $out -Encoding UTF8
    Write-Log "Tree [$Label] -> $out" OK
}

#endregion

# =====================================================================================
#region STORAGE / TRANSPORT  (OneDrive + F: unified via -Target)
# =====================================================================================

function Save-Backup {
    param([Parameter(Mandatory)][ValidateSet('OneDrive','FDrive')][string]$Target)
    $root = $Script:Config.Root
    if (-not (Test-Path -LiteralPath $root)) { Write-Log "Nothing to back up - $root missing." ERROR; return $false }

    switch ($Target) {
        'OneDrive' {
            $base = Test-OneDriveReady
            if (-not $base) { Write-Log 'OneDrive unavailable; skipping OneDrive backup.' WARN; return $false }
            $dest = Join-Path $base ("Backup_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            return (Invoke-Robocopy -Source $root -Destination (Join-Path $dest 'LaptopTransferBackups'))
        }
        'FDrive' {
            $base = Resolve-FDrive
            if (-not $base) { Write-Log 'F: share unavailable; skipping F: backup.' WARN; return $false }
            $dest = Join-Path $base "Backup_$env:USERNAME"
            # exclude bulky/irrelevant dirs on the network share (proper /XD, fixes old -like bug)
            return (Invoke-Robocopy -Source $root -Destination (Join-Path $dest 'LaptopTransferBackups') `
                    -Options @('/E','/COPY:DAT','/R:1','/W:1','/NP','/NFL','/NDL','/XD', $Script:Paths.Downloads, $Script:Paths.Wallpaper))
        }
    }
}

function Get-Backup {
    param([Parameter(Mandatory)][ValidateSet('OneDrive','FDrive')][string]$Target)
    $root = $Script:Config.Root
    New-DirIfMissing $root

    switch ($Target) {
        'OneDrive' {
            $base = Test-OneDriveReady
            if (-not $base) { Write-Log 'OneDrive unavailable.' ERROR; return $false }
            $backups = Get-ChildItem $base -Directory -ErrorAction SilentlyContinue |
                       Where-Object Name -match '^Backup_\d{8}_\d{6}$' | Sort-Object Name -Descending
            if (-not $backups) { Write-Log "No OneDrive backups in $base." ERROR; return $false }
            $pick = Select-FromList -Items $backups -Display { param($b) $b.Name } -Prompt 'Select a backup (Enter = latest)'
            if (-not $pick) { return $false }
            $inner = Join-Path $pick.FullName 'LaptopTransferBackups'
            $src   = if (Test-Path -LiteralPath $inner) { $inner } else { $pick.FullName }
            Write-Log "Restoring from $src" INFO
            return (Invoke-Robocopy -Source $src -Destination $root)
        }
        'FDrive' {
            $base = Resolve-FDrive
            if (-not $base) { Write-Log 'F: share unavailable.' ERROR; return $false }
            $folder = Join-Path $base "Backup_$env:USERNAME"
            if (-not (Test-Path -LiteralPath $folder)) { Write-Log "No F: backup for $env:USERNAME at $folder." ERROR; return $false }
            $inner = Join-Path $folder 'LaptopTransferBackups'
            $src   = if (Test-Path -LiteralPath $inner) { $inner } else { $folder }
            Write-Log "Restoring from $src" INFO
            return (Invoke-Robocopy -Source $src -Destination $root)
        }
    }
}

#endregion

# =====================================================================================
#region PIPELINES  (one backup + one parameterized restore)
# =====================================================================================

function Invoke-SwapStep {
    # Run one pipeline step, record its result, and verify the artifact it should produce.
    # Status: Pass / Fail / Skip / Warn. A failure never aborts the run.
    #   Action : return $false  -> Fail (or Skip if -Optional); or the string 'SKIP' / 'WARN'.
    #   Verify : optional post-check; return $false -> Warn, or a string used as the detail text.
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Action,
        [scriptblock]$Verify,
        [switch]$Optional
    )
    $Script:StepIndex++
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log ("[{0}/{1}] {2}" -f $Script:StepIndex, $Script:StepTotal, $Name) STEP

    $status = 'Pass'; $detail = ''
    try {
        $r = & $Action
        if     ($r -is [string] -and $r -eq 'SKIP') { $status = 'Skip' }
        elseif ($r -is [string] -and $r -eq 'WARN') { $status = 'Warn' }
        elseif ($r -is [bool]   -and -not $r)       { $status = if ($Optional) { 'Skip' } else { 'Fail' } }
    } catch {
        $status = 'Fail'; $detail = $_.Exception.Message
    }

    if ($status -eq 'Pass' -and $Verify) {
        try {
            $v = & $Verify
            if     ($v -is [string])            { $detail = $v }
            elseif ($v -is [bool] -and -not $v) { $status = 'Warn'; $detail = 'output not verified' }
        } catch { $status = 'Warn'; $detail = "verify error: $($_.Exception.Message)" }
    }

    $sw.Stop()
    $Script:Steps.Add([pscustomobject]@{
        Step = $Name; Status = $status; Detail = $detail; Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
    })

    $line = "   {0} {1}" -f $Script:Marks[$status], $Name
    if ($detail) { $line += " ($detail)" }
    Write-Host $line -ForegroundColor $Script:MarkColor[$status]
}

function Show-Preflight {
    # Pre-run summary of the environment; returns $true if the tech confirms.
    param([Parameter(Mandatory)][string]$Mode)
    $od = Resolve-OneDrive
    $fd = Resolve-FDrive
    $admin = Test-Admin
    Write-Host ''
    Write-Host "----------------- PRE-FLIGHT ($Mode) -----------------" -ForegroundColor Cyan
    Write-Host ("   User            : {0}" -f $env:USERNAME)
    Write-Host ("   Computer        : {0}" -f $env:COMPUTERNAME)
    Write-Host ("   Admin rights    : {0}" -f $(if ($admin) { 'Yes' } else { 'NO - registry steps may fail' })) -ForegroundColor $(if ($admin) { 'Gray' } else { 'Yellow' })
    Write-Host ("   Office version  : {0}" -f $Script:OfficeVer)
    Write-Host ("   OneDrive        : {0}" -f $(if ($od) { $od } else { 'NOT found / not signed in' })) -ForegroundColor $(if ($od) { 'Gray' } else { 'Yellow' })
    Write-Host ("   F: backup share : {0}" -f $(if ($fd) { $fd } else { 'not present (non-divisional laptop)' }))
    Write-Host ("   Working folder  : {0}" -f $Script:Config.Root)
    Write-Host '------------------------------------------------------' -ForegroundColor Cyan
    return (Confirm-Action "Proceed with $Mode?" $true)
}

function Show-Checklist {
    param([Parameter(Mandatory)][string]$Title)
    Set-ConsoleUtf8
    $elapsed = if ($Script:RunStart) { (Get-Date) - $Script:RunStart } else { [TimeSpan]::Zero }
    $pass = @($Script:Steps | Where-Object Status -eq 'Pass').Count
    $fail = @($Script:Steps | Where-Object Status -eq 'Fail').Count
    $skip = @($Script:Steps | Where-Object Status -eq 'Skip').Count
    $warn = @($Script:Steps | Where-Object Status -eq 'Warn').Count

    Write-Host ''
    Write-Host ("================= {0} =================" -f $Title) -ForegroundColor Cyan
    foreach ($s in $Script:Steps) {
        $line = "  {0}  {1}" -f $Script:Marks[$s.Status], $s.Step
        if ($s.Detail) { $line += "  ($($s.Detail))" }
        $line += "  [{0}s]" -f $s.Seconds
        Write-Host $line -ForegroundColor $Script:MarkColor[$s.Status]
    }
    Write-Host ('-' * 54) -ForegroundColor DarkGray
    $summary = "  {0} passed" -f $pass
    if ($fail) { $summary += " - {0} failed"   -f $fail }
    if ($warn) { $summary += " - {0} warnings" -f $warn }
    if ($skip) { $summary += " - {0} skipped"  -f $skip }
    $summary += "        Elapsed {0:mm\:ss}" -f $elapsed
    $sumColor = if ($fail) { 'Red' } elseif ($warn) { 'Yellow' } else { 'Green' }
    Write-Host $summary -ForegroundColor $sumColor
    if ($Script:LogFile) { Write-Host ("  Log: {0}" -f $Script:LogFile) -ForegroundColor DarkGray }
    Write-Host ('=' * (38 + $Title.Length)) -ForegroundColor Cyan
}

function Invoke-Pipeline {
    # Runs an ordered list of step definitions, then prints the sanity checklist.
    #   $Steps : array of hashtables @{ Name; Action; Verify (optional); Optional (optional) }
    param([Parameter(Mandatory)][object[]]$Steps, [Parameter(Mandatory)][string]$Title)
    $Script:Steps     = New-Object System.Collections.Generic.List[object]
    $Script:StepIndex = 0
    $Script:StepTotal = $Steps.Count
    $Script:RunStart  = Get-Date
    foreach ($s in $Steps) {
        Invoke-SwapStep -Name $s.Name -Action $s.Action -Verify $s.Verify -Optional:([bool]$s.Optional)
    }
    Show-Checklist -Title $Title
}

function Start-SwapBackup {
    Set-ConsoleUtf8
    if (-not (Show-Preflight -Mode 'BACKUP')) { Write-Log 'Backup cancelled by user.' WARN; return }
    Clear-StaleWorkingFolder
    New-DirIfMissing $Script:Config.Root
    Start-SwapLog
    try {
        $steps = @(
            @{ Name='Device summary';        Action={ Get-DeviceSummary };                          Verify={ Test-Path $Script:Paths.ComputerInfo } }
            @{ Name='Email signatures';      Action={ Sync-Signature -Direction Backup }; Optional=$true; Verify={ "$(@(Get-ChildItem $Script:Paths.Signatures -Recurse -File -ErrorAction SilentlyContinue).Count) files" } }
            @{ Name='OneNote notebook list'; Action={ Export-OneNoteNotebooks };                    Verify={ $j=Read-JsonFile $Script:Paths.OneNoteJson; if ($j) { "$(@($j).Count) notebooks" } else { $false } } }
            @{ Name='OneNote shortcuts';     Action={ New-OneNoteShortcuts }; Optional=$true;       Verify={ Test-Path $Script:Paths.OneNoteShortcuts } }
            @{ Name='OneNote registry';      Action={ Export-OneNoteRegistry };                     Verify={ Test-Path $Script:Paths.OneNoteReg } }
            @{ Name='Outlook profile';       Action={ Export-OutlookProfile -Name 'OldPcOutlook' }; Verify={ Test-Path (Join-Path $Script:Paths.OutlookRegDir 'OldPcOutlook.reg') } }
            @{ Name='Quick Access';          Action={ Backup-QuickAccess };                         Verify={ $n=@(Get-ChildItem $Script:Paths.QuickAccess -File -ErrorAction SilentlyContinue).Count; if ($n -gt 0) { "$n files" } else { $false } } }
            @{ Name='Downloads';             Action={ Backup-Downloads }; Optional=$true;           Verify={ @(Get-ChildItem $Script:Paths.Downloads -Filter *.zip -ErrorAction SilentlyContinue).Count -gt 0 } }
            @{ Name='Wallpaper';             Action={ Backup-Wallpaper }; Optional=$true;           Verify={ @(Get-ChildItem $Script:Paths.Wallpaper -File -ErrorAction SilentlyContinue).Count -gt 0 } }
            @{ Name='Folder trees';          Action={ Save-FolderTree -Directory (Join-Path $env:USERPROFILE 'Downloads') -Label 'Downloads' }; Verify={ Test-Path (Join-Path $Script:Paths.Trees 'Downloads.txt') } }
            @{ Name='App inventory';         Action={ Get-InstalledAppList };                       Verify={ Test-Path $Script:Paths.AppList } }
            @{ Name='Push to OneDrive';      Action={ Save-Backup -Target OneDrive }; Optional=$true }
            @{ Name='Push to F: drive';      Action={ Save-Backup -Target FDrive };   Optional=$true }
        )
        Invoke-Pipeline -Steps $steps -Title 'BACKUP SUMMARY'
    } finally { Stop-SwapLog }
}

function Start-SwapRestore {
    param(
        [ValidateSet('OneDrive','FDrive','Offline')][string]$Source = 'OneDrive',
        [bool]$IncludeOutlook = $true
    )
    Set-ConsoleUtf8
    if (-not (Show-Preflight -Mode "RESTORE ($Source)")) { Write-Log 'Restore cancelled by user.' WARN; return }
    New-DirIfMissing $Script:Config.Root
    Start-SwapLog
    try {
        $steps = New-Object System.Collections.Generic.List[object]

        if ($IncludeOutlook) {
            $steps.Add(@{ Name='Launch Outlook (first run)'; Optional=$true;
                          Action={ $exe = Resolve-OutlookExe; if ($exe) { Start-Process $exe } else { 'SKIP' } } })
        }

        switch ($Source) {
            'OneDrive' { $steps.Add(@{ Name='Fetch backup (OneDrive)'; Action={ Get-Backup -Target OneDrive }; Verify={ @(Get-ChildItem $Script:Config.Root -ErrorAction SilentlyContinue).Count -gt 0 } }) }
            'FDrive'   { $steps.Add(@{ Name='Fetch backup (F: drive)'; Action={ Get-Backup -Target FDrive };   Verify={ @(Get-ChildItem $Script:Config.Root -ErrorAction SilentlyContinue).Count -gt 0 } }) }
            'Offline'  { $steps.Add(@{ Name='Stage offline files'; Action={
                            Write-Log "Offline restore: put the backup files under $($Script:Config.Root) (structure: <Root>\*files*)." WARN
                            Write-Log "Download from $($Script:Config.SharePointUrl) if needed." INFO
                            if (-not $Script:Unattended) { Read-Host 'Press Enter once the files are in place' | Out-Null }
                         } }) }
        }

        $steps.Add(@{ Name='Restore signatures';   Action={ Sync-Signature -Direction Restore }; Optional=$true; Verify={ Test-Path (Join-Path $env:APPDATA 'Microsoft\Signatures') } })
        $steps.Add(@{ Name='Restore Quick Access'; Action={ Restore-QuickAccess }; Verify={ @(Get-ChildItem (Join-Path $env:APPDATA 'Microsoft\Windows\Recent\AutomaticDestinations') -File -ErrorAction SilentlyContinue).Count -gt 0 } })
        $steps.Add(@{ Name='OneNote shortcuts';    Action={ New-OneNoteShortcuts }; Optional=$true; Verify={ Test-Path $Script:Paths.OneNoteShortcuts } })
        $steps.Add(@{ Name='Compare notebooks';    Action={ Compare-OneNoteNotebook } })
        $steps.Add(@{ Name='Import OneNote reg';   Action={ Import-OneNoteRegistry }; Verify={ Test-Path "HKCU:\Software\Microsoft\Office\$Script:OfficeVer\OneNote\OpenNotebooks" } })

        if ($IncludeOutlook) {
            $steps.Add(@{ Name='Wait for Outlook profile'; Action={ Wait-ForOutlookProfile }; Verify={ Test-Path $Script:Reg.OutlookPS } })
            $steps.Add(@{ Name='Import Outlook profile';   Action={ Import-OutlookProfile -Name 'OldPcOutlook' }; Verify={ Test-Path $Script:Reg.OutlookPS } })
        }

        Invoke-Pipeline -Steps $steps.ToArray() -Title 'RESTORE SUMMARY'
    } finally { Stop-SwapLog }
}

#endregion

# =====================================================================================
#region MENU / ENTRY POINT
# =====================================================================================

function Show-Menu {
    $items = [ordered]@{
        '1' = @{ Group = 'Backup';  Text = 'Full Backup';                        Action = { Start-SwapBackup } }
        '2' = @{ Group = 'Restore'; Text = 'Restore from OneDrive';              Action = { Start-SwapRestore -Source OneDrive -IncludeOutlook $true } }
        '3' = @{ Group = 'Restore'; Text = 'Restore from F: drive';              Action = { Start-SwapRestore -Source FDrive  -IncludeOutlook $true } }
        '4' = @{ Group = 'Restore'; Text = 'Restore from OneDrive (no Outlook)'; Action = { Start-SwapRestore -Source OneDrive -IncludeOutlook $false } }
        '5' = @{ Group = 'Restore'; Text = 'Offline Restore (manual files)';     Action = { Start-SwapRestore -Source Offline -IncludeOutlook $false } }
        '6' = @{ Group = 'Tools';   Text = 'Repair Outlook profile';             Action = { Repair-OutlookProfile } }
        '7' = @{ Group = 'Tools';   Text = 'Show installed apps';                Action = { Get-InstalledAppList } }
        '8' = @{ Group = 'Tools';   Text = 'Check OneDrive connection';          Action = { $od = Resolve-OneDrive; if ($od) { Write-Log "OneDrive OK: $od" OK } else { Write-Log 'OneDrive not found / not signed in.' WARN } } }
    }

    while ($true) {
        Clear-Host
        $admin = Test-Admin
        $badge = if ($admin) { '[ADMIN]' } else { '[NOT ADMIN]' }
        Write-Host '==================== ALDI LAPTOP SWAP (v21) ====================' -ForegroundColor Cyan
        Write-Host ("  {0}  on  {1}    {2}" -f $env:USERNAME, $env:COMPUTERNAME, $badge) -ForegroundColor $(if ($admin) { 'Green' } else { 'Yellow' })
        if (-not $admin) { Write-Host '  (registry import/export may fail without admin rights)' -ForegroundColor Yellow }

        $lastGroup = $null
        foreach ($k in $items.Keys) {
            if ($items[$k].Group -ne $lastGroup) {
                Write-Host ''
                Write-Host ("  -- {0} --" -f $items[$k].Group) -ForegroundColor DarkCyan
                $lastGroup = $items[$k].Group
            }
            Write-Host ("    {0}. {1}" -f $k, $items[$k].Text)
        }
        Write-Host ''
        Write-Host '    Q. Quit'
        Write-Host ''

        $sel = (Read-Host ' Select an option').Trim().ToUpper()
        if ($sel -eq 'Q') { break }
        if ($items.Contains($sel)) {
            try { & $items[$sel].Action }
            catch { Write-Log "Action failed: $($_.Exception.Message)" ERROR }
            Write-Host ''
            Write-Host ' Press any key to return to the menu...' -ForegroundColor DarkGray
            try { [void][System.Console]::ReadKey($true) } catch { Read-Host | Out-Null }
        } else {
            Write-Host ("Invalid selection: '{0}'" -f $sel) -ForegroundColor Yellow
            Start-Sleep -Milliseconds 900
        }
    }
}

# ---- entry point -------------------------------------------------------------------
if (-not $Import) {
    Set-ConsoleUtf8
    try {
        Show-Menu
    } catch {
        Write-Log "Fatal error: $($_.Exception.Message)" ERROR
    } finally {
        Stop-SwapLog
    }
}

#endregion
