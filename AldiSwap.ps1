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
      * Late-bound OneNote COM + auto-detected Office version (no fragile GAC/2013 schema).
      * Edge/Chrome bookmarks captured; a readable manifest.json travels with each backup.
      * Live health strip (OneDrive/F:/network), per-step progress, and a sanity dashboard.
      * Flags local OneNote notebooks (won't migrate) and any that don't reopen on restore.
      * Approved verbs, StrictMode, UTF-8 everywhere, exit-code checks. No admin required.

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

Set-StrictMode -Version Latest      # surfaces typos / unset variables instead of failing silently
$ErrorActionPreference = 'Stop'     # cmdlet errors throw and are caught per-step by the pipeline

function Get-OfficeVersion {
    # Highest installed Office version key (e.g. 16.0) that actually contains Outlook/OneNote.
    # Replaces the old hardcoded '16.0' so registry paths follow the machine, not an assumption.
    try {
        $root = 'HKCU:\Software\Microsoft\Office'
        if (Test-Path $root) {
            $hit = Get-ChildItem $root -ErrorAction SilentlyContinue |
                   Where-Object { $_.PSChildName -match '^\d{2}\.\d$' } |
                   Sort-Object { [double]$_.PSChildName } -Descending |
                   Where-Object { (Test-Path (Join-Path $_.PSPath 'Outlook')) -or (Test-Path (Join-Path $_.PSPath 'OneNote')) } |
                   Select-Object -First 1
            if ($hit) { return $hit.PSChildName }
        }
    } catch { }
    return '16.0'
}

# =====================================================================================
#region CONFIG  -- single source of truth (edit values here, nowhere else)
# =====================================================================================

$Script:OfficeVer = Get-OfficeVersion   # auto-detected Office major version for registry paths

$Script:Config = [ordered]@{
    Root          = 'C:\Temp\LaptopTransferBackups'                       # local working folder
    FDriveRoot    = 'F:\usrnew\For IT Support\Laptop Swap Script\Backups' # divisional network backup
    SharePointUrl = 'https://asgportal-my.sharepoint.com/my'             # offline-restore download source
    CorpHost      = '10.60.162.200'                                       # network reachability probe (DNS/server)
    PingTimeoutMs = 1000                                                  # bounded so the health check never hangs
    # OneDrive folders are "OneDrive - <tenant>"; tried in this order (most specific first):
    OneDriveTenants = @('OneDrive - ALDI DX', 'OneDrive - ALDI-HOFER')
    Processes     = @{ OneNote = 'onenote'; OneNoteSender = 'ONENOTEM'; Outlook = 'OUTLOOK'; Explorer = 'explorer'; OneDrive = 'OneDrive' }
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
    Bookmarks          = Join-Path $Script:Config.Root 'BrowserBookmarks'
    Manifest           = Join-Path $Script:Config.Root 'manifest.json'
}

# Registry keys (reg.exe form for export/import; PS-provider form for Test-Path)
$Script:Reg = @{
    OneNoteOpenExe = "HKEY_CURRENT_USER\Software\Microsoft\Office\$Script:OfficeVer\OneNote\OpenNotebooks"
    OutlookExe     = "HKEY_CURRENT_USER\Software\Microsoft\Office\$Script:OfficeVer\Outlook\Profiles"
    OutlookPS      = "HKCU:\Software\Microsoft\Office\$Script:OfficeVer\Outlook\Profiles"
    WallpaperPS    = 'HKCU:\Control Panel\Desktop'
}

# Browsers whose Bookmarks file we capture/restore (profiles discovered at runtime).
$Script:Browsers = @(
    @{ Name = 'Edge';   Base = 'Microsoft\Edge\User Data'; Process = 'msedge' }
    @{ Name = 'Chrome'; Base = 'Google\Chrome\User Data';  Process = 'chrome' }
)

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

# Cached connectivity/health + backup counts (populated by Update-HealthStatus)
$Script:Health       = $null
$Script:BackupCounts = @{ F = 0; O = 0 }

# Sanity-check / progress state (populated per pipeline run)
$Script:Steps       = New-Object System.Collections.Generic.List[object]
$Script:StepIndex   = 0
$Script:StepTotal   = 0
$Script:RunStart    = $null
$Script:ReportFlags = New-Object System.Collections.Generic.List[object]  # end-of-run "Attention" items

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
    $sw   = [System.Diagnostics.Stopwatch]::StartNew()
    $spin = [char[]]'|/-\'
    $i    = 0
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (& $Condition) {
            if ($i -gt 0) { Write-Host ("`r   {0,-60}" -f "$Activity... done.") -ForegroundColor DarkGray }
            return $true
        }
        $remain = [int]($TimeoutSeconds - $sw.Elapsed.TotalSeconds)
        Write-Host ("`r   {0} {1}... {2,3}s left " -f $spin[$i % 4], $Activity, $remain) -NoNewline -ForegroundColor DarkGray
        $i++
        Start-Sleep -Seconds $PollSeconds
    }
    Write-Host ("`r   {0,-60}" -f "$Activity... timed out.")
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

function Reset-ReportFlags {
    $Script:ReportFlags = New-Object System.Collections.Generic.List[object]
}

function Add-ReportFlag {
    # Queue an item for the end-of-run "Attention" section (e.g. local notebooks, missing data).
    param(
        [ValidateSet('Warn','Error')][string]$Level = 'Warn',
        [Parameter(Mandatory)][string]$Title,
        [string]$Detail = ''
    )
    $Script:ReportFlags.Add([pscustomobject]@{ Level = $Level; Title = $Title; Detail = $Detail })
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
#region CONNECTIVITY / HEALTH
# =====================================================================================

function Test-CorpNetwork {
    # Fast, bounded reachability probe to the corp host (System.Net so it can't hang the UI).
    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        return (($ping.Send($Script:Config.CorpHost, $Script:Config.PingTimeoutMs)).Status -eq 'Success')
    } catch { return $false }
}

function Test-OneDriveRunning {
    return [bool](Get-Process -Name $Script:Config.Processes.OneDrive -ErrorAction SilentlyContinue)
}

function Update-HealthStatus {
    # Compute and CACHE health + backup counts so the menu header renders instantly without
    # re-probing the network on every keypress. Called at startup, on [R]efresh, and per run.
    # The F: probe is gated on network reachability so an offline box never hangs on a dead share.
    $corp   = Test-CorpNetwork
    $odPath = $null
    try { $odPath = Resolve-OneDrive } catch { }            # local folder test, fast
    $odRun  = Test-OneDriveRunning

    $fdPath = $null
    if ($corp) { try { $fdPath = Resolve-FDrive } catch { } }   # only touch the share when the network is up

    $oCount = 0
    if ($odPath) {
        try { $oCount = @(Get-ChildItem $odPath -Directory -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match '^Backup_\d{8}_\d{6}$' }).Count } catch { }
    }
    $fCount = 0
    if ($fdPath) {
        try { if (Test-Path -LiteralPath (Join-Path $fdPath "Backup_$env:USERNAME")) { $fCount = 1 } } catch { }
    }

    $Script:Health = [ordered]@{
        OneDriveFound   = [bool]$odPath
        OneDriveRunning = $odRun
        OneDriveOnline  = ([bool]$odPath -and $odRun -and $corp)
        FDriveOnline    = [bool]$fdPath
        CorpReachable   = $corp
        CheckedAt       = (Get-Date)
    }
    $Script:BackupCounts = @{ F = $fCount; O = $oCount }
    return $Script:Health
}

function Get-HealthColor { param([bool]$Ok) if ($Ok) { 'Green' } else { 'Red' } }

function Write-HealthItem {
    # One coloured status pill on the health strip. State: ok (green check) / warn (yellow !) / bad (red X).
    param([ValidateSet('ok','warn','bad')][string]$State, [string]$Label)
    switch ($State) {
        'ok'   { $mark = [char]0x2713; $col = 'Green'  }
        'warn' { $mark = [char]0x0021; $col = 'Yellow' }
        default{ $mark = [char]0x2717; $col = 'Red'    }
    }
    Write-Host ("  {0} {1}" -f $mark, $Label) -NoNewline -ForegroundColor $col
}

function Show-AppHeader {
    param([string]$Subtitle = '')
    Set-ConsoleUtf8
    if (-not $Script:Health) { Update-HealthStatus | Out-Null }
    $h = $Script:Health
    $c = $Script:BackupCounts

    Write-Host ''
    Write-Host '  ===============================================================' -ForegroundColor Cyan
    Write-Host '   ALDI LAPTOP SWAP TOOLKIT      v21' -ForegroundColor White
    Write-Host ("   {0} on {1}      Office {2}" -f $env:USERNAME, $env:COMPUTERNAME, $Script:OfficeVer) -ForegroundColor DarkGray
    if ($Subtitle) { Write-Host ("   >> {0}" -f $Subtitle) -ForegroundColor Cyan }
    Write-Host '  ---------------------------------------------------------------' -ForegroundColor DarkCyan

    # Health strip (each pill coloured independently)
    $odState = if ($h.OneDriveOnline) { 'ok' } elseif ($h.OneDriveFound) { 'warn' } else { 'bad' }
    $odLabel = if ($h.OneDriveOnline) { 'OneDrive online' }
               elseif ($h.OneDriveFound -and -not $h.OneDriveRunning) { 'OneDrive NOT running' }
               elseif ($h.OneDriveFound) { 'OneDrive offline' }
               else { 'OneDrive missing' }
    $fdState = if ($h.FDriveOnline) { 'ok' } else { 'bad' }
    $fdLabel = if ($h.FDriveOnline) { 'F: online' } else { 'F: unreachable' }
    $npState = if ($h.CorpReachable) { 'ok' } else { 'bad' }
    $npLabel = "Net $($Script:Config.CorpHost)"

    Write-Host '   Health ' -NoNewline -ForegroundColor Gray
    Write-HealthItem $odState $odLabel
    Write-HealthItem $fdState $fdLabel
    Write-HealthItem $npState $npLabel
    Write-Host ''

    # Backup counts (the "1F 2O" at-a-glance, item #5)
    Write-Host ("   Backups  {0}F  {1}O" -f $c.F, $c.O) -NoNewline -ForegroundColor White
    Write-Host ("   (F:/OneDrive found for {0})   checked {1:HH:mm:ss}" -f $env:USERNAME, $h.CheckedAt) -ForegroundColor DarkGray
    Write-Host '  ===============================================================' -ForegroundColor Cyan
}

function Show-HealthDetail {
    # Verbose connectivity report for the Tools menu item.
    Update-HealthStatus | Out-Null
    $h = $Script:Health
    Write-Host ''
    Write-Host '  Connectivity / health check' -ForegroundColor Cyan
    Write-Host '  ---------------------------------------------------------------' -ForegroundColor DarkCyan
    Write-Host ('   OneDrive folder  : {0}' -f $(if ($h.OneDriveFound) { Resolve-OneDrive } else { 'NOT found / not signed in' })) -ForegroundColor (Get-HealthColor $h.OneDriveFound)
    Write-Host ('   OneDrive.exe     : {0}' -f $(if ($h.OneDriveRunning) { 'running' } else { 'NOT running - files will not sync' })) -ForegroundColor (Get-HealthColor $h.OneDriveRunning)
    Write-Host ('   Network {0,-14}: {1}' -f $Script:Config.CorpHost, $(if ($h.CorpReachable) { 'reachable' } else { 'no response' })) -ForegroundColor (Get-HealthColor $h.CorpReachable)
    Write-Host ('   F: share         : {0}' -f $(if ($h.FDriveOnline) { Resolve-FDrive } else { 'unreachable' })) -ForegroundColor (Get-HealthColor $h.FDriveOnline)
    Write-Host ('   Backups found    : {0} on F:, {1} on OneDrive' -f $Script:BackupCounts.F, $Script:BackupCounts.O) -ForegroundColor White
    return $true
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
        # Late-bound COM: no GAC interop assembly, no version-specific type names to break on.
        $Script:OneNoteApp = New-Object -ComObject OneNote.Application
        return $Script:OneNoteApp
    } catch {
        Write-Log "Unable to create OneNote COM object (is OneNote installed and opened once?): $($_.Exception.Message)" ERROR
        return $null
    }
}

function Get-OneNoteNotebookList {
    $app = Get-OneNoteApp
    if (-not $app) { return @() }
    try {
        $xml = ''
        # 4 = hsPages: returns the full hierarchy; we only read the Notebook nodes out of it.
        $app.GetHierarchy([string]::Empty, 4, [ref]$xml) | Out-Null
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($xml)
        # local-name() ignores the schema-version namespace, so this works on any Office build.
        $doc.SelectNodes("//*[local-name()='Notebook']") | ForEach-Object {
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

function Test-NotebookIsCloud {
    # Cloud notebooks (SharePoint/OneDrive) reopen automatically from the restored registry.
    # Anything else (a C:\ path or a UNC share) is LOCAL and will NOT migrate during a swap.
    param([string]$Path)
    if (-not $Path) { return $false }
    return ($Path -match '^https?://')
}

function Find-LocalNotebook {
    # BACKUP-side check (#9): flag any captured notebook that isn't on SharePoint/OneDrive, so the
    # tech knows those won't come across the swap and must be handled by hand.
    $list = @(Read-JsonFile $Script:Paths.OneNoteJson)
    if (-not $list -or $list.Count -eq 0) { Write-Log 'No notebook list to inspect.' INFO; return $true }
    $local = @($list | Where-Object { -not (Test-NotebookIsCloud $_.Path) })
    if ($local.Count -eq 0) { Write-Log 'All notebooks are cloud (SharePoint/OneDrive).' OK; return $true }
    foreach ($nb in $local) {
        Add-ReportFlag -Level Warn -Title ("LOCAL notebook (won't migrate): {0}" -f $nb.Name) -Detail $nb.Path
        Write-Log (" - local notebook: {0} [{1}]" -f $nb.Name, $nb.Path) WARN
    }
    return 'WARN'
}

function Test-OneNoteRestore {
    # RESTORE-side check (#7): after importing the OpenNotebooks registry, open OneNote and let the
    # backed-up notebooks reopen, then report any that did NOT come back. Robust replacement for the
    # old fixed-Start-Sleep compare: polls until the live list stops growing (or every expected
    # notebook is present), so slow OneNote sync no longer produces false "missing" warnings.
    $expected = @(Read-JsonFile $Script:Paths.OneNoteJson)
    if (-not $expected -or $expected.Count -eq 0) { Write-Log 'No backed-up notebook list to verify against.' WARN; return 'WARN' }

    if (-not (Get-Process -Name $Script:Config.Processes.OneNote -ErrorAction SilentlyContinue)) {
        Start-Process 'onenote.exe' -ErrorAction SilentlyContinue
    }
    Wait-ForCondition -Condition { Get-Process -Name $Script:Config.Processes.OneNote -ErrorAction SilentlyContinue } `
                      -TimeoutSeconds 60 -PollSeconds 2 -Activity 'Opening OneNote' | Out-Null

    $expectedNames = @($expected | ForEach-Object { $_.Name })
    $deadline  = (Get-Date).AddSeconds(180)   # notebooks re-sync over time on a fresh PC
    $present   = @()
    $lastCount = -1
    $stableFor = 0

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 5
        $present    = @(Get-OneNoteNotebookList | ForEach-Object { $_.Name })
        $missingNow = @($expectedNames | Where-Object { $present -notcontains $_ })
        Write-Host ("`r   Waiting for notebooks to sync... {0}/{1} back   " -f ($expectedNames.Count - $missingNow.Count), $expectedNames.Count) -NoNewline -ForegroundColor DarkGray
        if ($missingNow.Count -eq 0) { break }
        if ($present.Count -eq $lastCount) { $stableFor++ } else { $stableFor = 0 }
        $lastCount = $present.Count
        if ($stableFor -ge 3) { break }   # list stable for ~15s -> stop waiting
    }
    Write-Host ''

    Write-JsonFile -Object (Get-OneNoteNotebookList) -Path $Script:Paths.OneNoteCompareJson
    $missing = @($expected | Where-Object { $present -notcontains $_.Name })
    if ($missing.Count -eq 0) { Write-Log ("All {0} notebook(s) restored." -f $expected.Count) OK; return $true }

    foreach ($nb in $missing) {
        Add-ReportFlag -Level Warn -Title ("Notebook NOT restored: {0}" -f $nb.Name) -Detail $nb.Path
        Write-Log (" - not restored: {0} [{1}]" -f $nb.Name, $nb.Path) WARN
    }
    return 'WARN'
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
    # Explorer is the Windows shell and AUTO-RESTARTS the instant it's killed, so we must NOT
    # wait for it to stay closed - that produced the bogus "Closing explorer timed out after 30s"
    # warning even though Explorer was perfectly fine. Just bounce it and confirm one is back.
    Write-Log 'Refreshing Explorer to pick up restored Quick Access...' INFO
    Get-Process -Name $Script:Config.Processes.Explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    if (-not (Get-Process -Name $Script:Config.Processes.Explorer -ErrorAction SilentlyContinue)) { Start-Process explorer.exe }
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

function Backup-BrowserBookmark {
    # Copy the Bookmarks JSON (and .bak) for every Edge/Chrome profile that has one.
    $root  = $Script:Paths.Bookmarks
    $count = 0
    $names = @()
    foreach ($b in $Script:Browsers) {
        $userData = Join-Path $env:LOCALAPPDATA $b.Base
        if (-not (Test-Path -LiteralPath $userData)) { continue }
        $names += $b.Name
        $profiles = @(Get-ChildItem $userData -Directory -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' })
        foreach ($p in $profiles) {
            $bm = Join-Path $p.FullName 'Bookmarks'
            if (-not (Test-Path -LiteralPath $bm)) { continue }
            $dst = Join-Path $root (Join-Path $b.Name $p.Name)
            New-DirIfMissing $dst
            Copy-Item -LiteralPath $bm -Destination $dst -Force
            if (Test-Path -LiteralPath "$bm.bak") { Copy-Item -LiteralPath "$bm.bak" -Destination $dst -Force }
            $count++
        }
    }
    if ($count -eq 0) { Write-Log 'No Edge/Chrome bookmarks found.' WARN; return $false }
    Write-Log ("Backed up bookmarks: {0} profile(s) from {1}" -f $count, ($names -join ', ')) OK
    return $true
}

function Restore-BrowserBookmark {
    $root = $Script:Paths.Bookmarks
    if (-not (Test-Path -LiteralPath $root)) { Write-Log 'No bookmarks in backup.' WARN; return $false }
    $count = 0
    foreach ($b in $Script:Browsers) {
        $src = Join-Path $root $b.Name
        if (-not (Test-Path -LiteralPath $src)) { continue }
        if (Get-Process -Name $b.Process -ErrorAction SilentlyContinue) {
            Write-Log "$($b.Name) is running - close it so restored bookmarks are not overwritten." WARN
        }
        $userData = Join-Path $env:LOCALAPPDATA $b.Base
        foreach ($pd in (Get-ChildItem $src -Directory -ErrorAction SilentlyContinue)) {
            $bm = Join-Path $pd.FullName 'Bookmarks'
            if (-not (Test-Path -LiteralPath $bm)) { continue }
            $dst = Join-Path $userData $pd.Name
            New-DirIfMissing $dst
            Copy-Item -LiteralPath $bm -Destination $dst -Force
            $count++
        }
    }
    if ($count -eq 0) { Write-Log 'No matching browser profiles to restore into.' WARN; return $false }
    Write-Log ("Restored bookmarks into {0} profile(s)." -f $count) OK
    return $true
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
        [PSCustomObject]@{
            Application = $name
            Installed   = [bool]$hit
            Path        = if ($hit) { $hit.FullName } else { 'not found' }
        }
    }

    # Print a coloured table of EVERY tracked app to the terminal (green check = installed).
    Write-Host ''
    Write-Host '  Installed application report' -ForegroundColor Cyan
    Write-Host '  ---------------------------------------------------------------' -ForegroundColor DarkCyan
    foreach ($r in $rows) {
        $mark = if ($r.Installed) { [char]0x2713 } else { [char]0x2717 }
        $col  = if ($r.Installed) { 'Green' } else { 'DarkGray' }
        Write-Host ("   {0}  {1,-26}  {2}" -f $mark, $r.Application, $r.Path) -ForegroundColor $col
    }
    $yes = @($rows | Where-Object Installed).Count
    Write-Host '  ---------------------------------------------------------------' -ForegroundColor DarkCyan
    Write-Host ("   {0} of {1} tracked apps installed" -f $yes, $rows.Count) -ForegroundColor White

    Write-JsonFile -Object $rows -Path $Script:Paths.AppList
    Write-Log "App inventory -> $($Script:Paths.AppList)" OK
    return $true
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

function Write-BackupManifest {
    # A readable record of the run that travels inside the backup (Restore shows it as
    # provenance; it also doubles as a second source of truth for the sanity check).
    $manifest = [ordered]@{
        Tool       = 'ALDI Laptop Swap'
        Version    = 'v21'
        User       = $env:USERNAME
        Computer   = $env:COMPUTERNAME
        Model      = (Get-CimInstance Win32_ComputerSystem).Model
        ServiceTag = (Get-CimInstance Win32_BIOS).SerialNumber
        OfficeVer  = $Script:OfficeVer
        CreatedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
        Steps      = @($Script:Steps | ForEach-Object { [ordered]@{ Step = $_.Step; Status = $_.Status; Detail = $_.Detail } })
    }
    Write-JsonFile -Object $manifest -Path $Script:Paths.Manifest
    Write-Log "Manifest written -> $($Script:Paths.Manifest)" OK
    return $true
}

function Save-Backup {
    param([Parameter(Mandatory)][ValidateSet('OneDrive','FDrive')][string]$Target)
    $root = $Script:Config.Root
    if (-not (Test-Path -LiteralPath $root)) { Write-Log "Nothing to back up - $root missing." ERROR; return $false }

    switch ($Target) {
        'OneDrive' {
            $base = Test-OneDriveReady
            if (-not $base) { Write-Log 'OneDrive folder not found - cannot back up to OneDrive.' ERROR; return $false }
            $dest = Join-Path $base ("Backup_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            if (-not (Invoke-Robocopy -Source $root -Destination (Join-Path $dest 'LaptopTransferBackups'))) { return $false }

            # The copy only landed in the LOCAL OneDrive folder - the real upload is OneDrive's
            # background sync, so confirm it can actually happen (items #1 and #2).
            if (-not (Test-OneDriveRunning)) {
                Write-Log 'OneDrive.exe is not running - files are staged locally but will NOT sync until it starts.' WARN
                return 'WARN'
            }
            if (-not (Test-CorpNetwork)) {
                Write-Log 'No network connectivity - OneDrive cannot upload this backup right now.' ERROR
                return $false   # offline = upload will not happen -> RED
            }
            Write-Log 'OneDrive backup staged and syncing.' OK
            return $true
        }
        'FDrive' {
            # F: is always mapped on these machines, so an unreachable F: is a genuine RED error
            # (something is wrong with the machine, the network, or the VPN).
            $base = Resolve-FDrive
            if (-not $base) { Write-Log 'F: drive unreachable - check the network/VPN or the drive mapping.' ERROR; return $false }
            $dest = Join-Path $base "Backup_$env:USERNAME"
            # exclude bulky/irrelevant dirs on the network share (proper /XD, fixes old -like bug)
            if (-not (Invoke-Robocopy -Source $root -Destination (Join-Path $dest 'LaptopTransferBackups') `
                    -Options @('/E','/COPY:DAT','/R:1','/W:1','/NP','/NFL','/NDL','/XD', $Script:Paths.Downloads, $Script:Paths.Wallpaper))) { return $false }
            Write-Log "F: backup written to $dest" OK
            return $true
        }
    }
}

function Select-RestoreBackup {
    # Build ONE list of available backups (F: first, then OneDrive newest-first), let the tech
    # pick (Enter / 0 = default = newest F: else newest OneDrive), and stage it into the working
    # folder. Also offers files already sitting in C:\Temp for an offline restore.
    # Returns the staged working-folder path, or $null if nothing was selected.
    $root    = $Script:Config.Root
    $entries = New-Object System.Collections.Generic.List[object]

    $fd = Resolve-FDrive
    if ($fd) {
        $fb = Join-Path $fd "Backup_$env:USERNAME"
        if (Test-Path -LiteralPath $fb) {
            $entries.Add([pscustomobject]@{ Source = 'F: drive'; Path = $fb; Date = (Get-Item $fb).LastWriteTime })
        }
    }
    $od = Resolve-OneDrive
    if ($od) {
        Get-ChildItem $od -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^Backup_\d{8}_\d{6}$' } | Sort-Object Name -Descending |
            ForEach-Object { $entries.Add([pscustomobject]@{ Source = 'OneDrive'; Path = $_.FullName; Date = $_.LastWriteTime }) }
    }

    $tempHasData = (Test-Path -LiteralPath $root) -and
                   (@(Get-ChildItem $root -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'Logs' }).Count -gt 0)

    if ($entries.Count -eq 0 -and -not $tempHasData) {
        Write-Log 'No backups found on F: or OneDrive, and the local Temp folder is empty.' ERROR
        Write-Log "Offline option: download a backup from $($Script:Config.SharePointUrl) into $root, then retry." INFO
        return $null
    }

    Write-Host ''
    Write-Host '  Available backups:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $entries.Count; $i++) {
        $e   = $entries[$i]
        $def = if ($i -eq 0) { '   <- default' } else { '' }
        Write-Host ("    {0}. {1,-9}  {2:yyyy-MM-dd HH:mm}  {3}{4}" -f $i, $e.Source, $e.Date, (Split-Path $e.Path -Leaf), $def)
    }
    if ($tempHasData) { Write-Host '    T. Use files already staged in C:\Temp (offline restore)' }

    if ($Script:Unattended) { $sel = '0' }
    else { $sel = (Read-Host '  Select a backup (Enter = default)').Trim().ToUpper() }

    if ($sel -eq 'T' -and $tempHasData) { Write-Log "Using files already staged in $root (offline restore)." INFO; return $root }
    if ($sel -eq '') { $sel = '0' }

    if ($entries.Count -eq 0) {
        if ($tempHasData) { return $root }
        return $null
    }
    if (-not ($sel -match '^\d+$') -or [int]$sel -ge $entries.Count) {
        if ($tempHasData) { Write-Log 'Invalid choice - using locally staged files.' WARN; return $root }
        Write-Log 'Invalid choice.' ERROR; return $null
    }

    $pick = $entries[[int]$sel]
    # Some backups nest their content under a 'LaptopTransferBackups' subfolder - unwrap it.
    $inner = Join-Path $pick.Path 'LaptopTransferBackups'
    $src   = if (Test-Path -LiteralPath $inner) { $inner } else { $pick.Path }
    Write-Log "Staging backup from $($pick.Source): $src" INFO
    New-DirIfMissing $root
    if (Invoke-Robocopy -Source $src -Destination $root) { return $root }
    return $null
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
    # Quiet when healthy (per design): prints a one-line OK and proceeds. Only flashes a
    # checklist when something is wrong, and only hard-stops a backup that has nowhere to go.
    param([Parameter(Mandatory)][string]$Mode)
    $od = Resolve-OneDrive
    $fd = Resolve-FDrive
    $issues = New-Object System.Collections.Generic.List[string]

    if ($Mode -like 'BACKUP*') {
        if (-not $od -and -not $fd) { $issues.Add('No backup target: OneDrive not signed in AND no F: drive present.') }
        elseif (-not $od)           { $issues.Add('OneDrive not signed in - backup will go to the F: drive only.') }
        $c = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
        if ($c -and (($c.FreeSpace / 1GB) -lt 5)) { $issues.Add(('Low free space on C: ({0} GB) - large Downloads may not zip.' -f [math]::Round($c.FreeSpace/1GB,1))) }
    }

    if ($issues.Count -eq 0) {
        Write-Host ("  Pre-flight OK   Office {0} | OneDrive: {1} | F: {2}" -f `
            $Script:OfficeVer, $(if ($od) { 'yes' } else { 'no' }), $(if ($fd) { 'yes' } else { 'n/a' })) -ForegroundColor Green
        return $true
    }

    Write-Host ''
    Write-Host "----------------- PRE-FLIGHT ($Mode) -----------------" -ForegroundColor Cyan
    foreach ($i in $issues) { Write-Host "   ! $i" -ForegroundColor Yellow; Write-Log "Preflight: $i" WARN }
    Write-Host '------------------------------------------------------' -ForegroundColor Cyan

    if ($Mode -like 'BACKUP*' -and -not $od -and -not $fd) {
        $launcher = Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'
        if ((Test-Path -LiteralPath $launcher) -and (Confirm-Action 'Try to launch OneDrive now?' $true)) { Start-Process $launcher }
        return (Confirm-Action 'No backup target available yet. Continue anyway?' $false)
    }
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

    $total = $Script:Steps.Count

    Write-Host ''
    Write-Host ("  ==============  {0}  ==============" -f $Title) -ForegroundColor Cyan
    foreach ($s in $Script:Steps) {
        Write-Host ("   {0}  {1,-26}" -f $Script:Marks[$s.Status], $s.Step) -NoNewline -ForegroundColor $Script:MarkColor[$s.Status]
        if ($s.Detail) { Write-Host (" {0}" -f $s.Detail) -NoNewline -ForegroundColor $Script:MarkColor[$s.Status] }
        Write-Host ("  [{0}s]" -f $s.Seconds) -ForegroundColor DarkGray
    }
    Write-Host ('  ' + ('-' * 60)) -ForegroundColor DarkGray

    # Headline - green only when nothing failed and nothing warned.
    $headColor = if ($fail) { 'Red' } elseif ($warn) { 'Yellow' } else { 'Green' }
    Write-Host ("   RESULT:  {0} / {1} passed     ({2} warn, {3} failed, {4} skipped)     Elapsed {5:mm\:ss}" -f `
        $pass, $total, $warn, $fail, $skip, $elapsed) -ForegroundColor $headColor

    # Attention section - flagged local notebooks, notebooks that didn't restore, etc.
    if ($Script:ReportFlags.Count -gt 0) {
        Write-Host ''
        Write-Host ("   !! ATTENTION - {0} item(s) need a human !!" -f $Script:ReportFlags.Count) -ForegroundColor Yellow
        foreach ($f in $Script:ReportFlags) {
            $col = if ($f.Level -eq 'Error') { 'Red' } else { 'Yellow' }
            Write-Host ("     - {0}" -f $f.Title) -ForegroundColor $col
            if ($f.Detail) { Write-Host ("         {0}" -f $f.Detail) -ForegroundColor DarkGray }
        }
    }

    if ($Script:LogFile) { Write-Host ''; Write-Host ("   Log: {0}" -f $Script:LogFile) -ForegroundColor DarkGray }
    Write-Host ('  ' + ('=' * 60)) -ForegroundColor Cyan
}

function Invoke-Pipeline {
    # Runs an ordered list of step definitions, then prints the sanity checklist.
    #   $Steps : array of hashtables @{ Name; Action; Verify (optional); Optional (optional) }
    param([Parameter(Mandatory)][object[]]$Steps, [Parameter(Mandatory)][string]$Title)
    $Script:Steps     = New-Object System.Collections.Generic.List[object]
    $Script:StepIndex = 0
    $Script:StepTotal = $Steps.Count
    $Script:RunStart  = Get-Date
    $n = 0
    foreach ($s in $Steps) {
        $n++
        Write-Progress -Activity $Title -Status ("[{0}/{1}] {2}" -f $n, $Steps.Count, $s.Name) -PercentComplete ([int]((($n - 1) / $Steps.Count) * 100))
        # Verify/Optional are optional keys; access via ContainsKey so StrictMode stays happy.
        $verify   = if ($s.ContainsKey('Verify'))   { $s.Verify }        else { $null }
        $optional = if ($s.ContainsKey('Optional')) { [bool]$s.Optional } else { $false }
        Invoke-SwapStep -Name $s.Name -Action $s.Action -Verify $verify -Optional:$optional
    }
    Write-Progress -Activity $Title -Completed
    Show-Checklist -Title $Title
}

function Start-SwapBackup {
    Set-ConsoleUtf8
    Clear-Host
    Update-HealthStatus | Out-Null
    Show-AppHeader -Subtitle 'FULL BACKUP'
    if (-not (Show-Preflight -Mode 'BACKUP')) { Write-Log 'Backup cancelled by user.' WARN; return }
    Clear-StaleWorkingFolder
    New-DirIfMissing $Script:Config.Root
    Reset-ReportFlags
    Start-SwapLog
    try {
        $steps = @(
            @{ Name='Device summary';        Action={ Get-DeviceSummary };                          Verify={ Test-Path $Script:Paths.ComputerInfo } }
            @{ Name='Email signatures';      Action={ Sync-Signature -Direction Backup }; Optional=$true; Verify={ "$(@(Get-ChildItem $Script:Paths.Signatures -Recurse -File -ErrorAction SilentlyContinue).Count) files" } }
            @{ Name='OneNote notebook list'; Action={ Export-OneNoteNotebooks };                    Verify={ $j=Read-JsonFile $Script:Paths.OneNoteJson; if ($j) { "$(@($j).Count) notebooks" } else { $false } } }
            @{ Name='Local-notebook check';  Action={ Find-LocalNotebook } }
            @{ Name='OneNote shortcuts';     Action={ New-OneNoteShortcuts }; Optional=$true;       Verify={ Test-Path $Script:Paths.OneNoteShortcuts } }
            @{ Name='OneNote registry';      Action={ Export-OneNoteRegistry };                     Verify={ Test-Path $Script:Paths.OneNoteReg } }
            @{ Name='Outlook profile';       Action={ Export-OutlookProfile -Name 'OldPcOutlook' }; Verify={ Test-Path (Join-Path $Script:Paths.OutlookRegDir 'OldPcOutlook.reg') } }
            @{ Name='Quick Access';          Action={ Backup-QuickAccess };                         Verify={ $n=@(Get-ChildItem $Script:Paths.QuickAccess -File -ErrorAction SilentlyContinue).Count; if ($n -gt 0) { "$n files" } else { $false } } }
            @{ Name='Downloads';             Action={ Backup-Downloads }; Optional=$true;           Verify={ @(Get-ChildItem $Script:Paths.Downloads -Filter *.zip -ErrorAction SilentlyContinue).Count -gt 0 } }
            @{ Name='Wallpaper';             Action={ Backup-Wallpaper }; Optional=$true;           Verify={ @(Get-ChildItem $Script:Paths.Wallpaper -File -ErrorAction SilentlyContinue).Count -gt 0 } }
            @{ Name='Browser bookmarks';     Action={ Backup-BrowserBookmark }; Optional=$true;     Verify={ Test-Path $Script:Paths.Bookmarks } }
            @{ Name='Folder trees';          Action={ Save-FolderTree -Directory (Join-Path $env:USERPROFILE 'Downloads') -Label 'Downloads' }; Verify={ Test-Path (Join-Path $Script:Paths.Trees 'Downloads.txt') } }
            @{ Name='App inventory';         Action={ Get-InstalledAppList };                       Verify={ Test-Path $Script:Paths.AppList } }
            @{ Name='Write manifest';        Action={ Write-BackupManifest };                       Verify={ Test-Path $Script:Paths.Manifest } }
            @{ Name='Push to OneDrive';      Action={ Save-Backup -Target OneDrive } }
            @{ Name='Push to F: drive';      Action={ Save-Backup -Target FDrive } }
        )
        Invoke-Pipeline -Steps $steps -Title 'BACKUP SUMMARY'
    } finally { Stop-SwapLog }
}

function Start-SwapRestore {
    Set-ConsoleUtf8
    Clear-Host
    Update-HealthStatus | Out-Null
    Show-AppHeader -Subtitle 'RESTORE'
    if (-not (Show-Preflight -Mode 'RESTORE')) { Write-Log 'Restore cancelled by user.' WARN; return }
    New-DirIfMissing $Script:Config.Root
    Reset-ReportFlags
    Start-SwapLog
    try {
        # Let the tech choose which backup to restore (F: / OneDrive / offline), then stage it.
        $staged = Select-RestoreBackup
        if (-not $staged) { Write-Log 'Restore cancelled - no backup selected.' WARN; return }

        # If the backup carries a manifest, show where it came from (guard props for StrictMode).
        $mf = Read-JsonFile $Script:Paths.Manifest
        if ($mf -and ($mf.PSObject.Properties.Name -contains 'Computer')) {
            Write-Log ("Restoring backup from {0} ({1}) created {2}" -f $mf.Computer, $mf.User, $mf.CreatedUtc) INFO
        }

        $steps = New-Object System.Collections.Generic.List[object]
        # Launch Outlook early so it can build its profile while the other steps run.
        $steps.Add(@{ Name='Launch Outlook (first run)'; Optional=$true;
                      Action={ $exe = Resolve-OutlookExe; if ($exe) { Start-Process $exe } else { 'SKIP' } } })
        $steps.Add(@{ Name='Restore signatures';        Action={ Sync-Signature -Direction Restore }; Optional=$true; Verify={ Test-Path (Join-Path $env:APPDATA 'Microsoft\Signatures') } })
        $steps.Add(@{ Name='Restore Quick Access';      Action={ Restore-QuickAccess }; Verify={ @(Get-ChildItem (Join-Path $env:APPDATA 'Microsoft\Windows\Recent\AutomaticDestinations') -File -ErrorAction SilentlyContinue).Count -gt 0 } })
        $steps.Add(@{ Name='Restore browser bookmarks'; Action={ Restore-BrowserBookmark }; Optional=$true })
        $steps.Add(@{ Name='OneNote shortcuts';         Action={ New-OneNoteShortcuts }; Optional=$true; Verify={ Test-Path $Script:Paths.OneNoteShortcuts } })
        # Import the OpenNotebooks registry FIRST, then open OneNote and verify they reopened.
        $steps.Add(@{ Name='Import OneNote reg';        Action={ Import-OneNoteRegistry }; Verify={ Test-Path "HKCU:\Software\Microsoft\Office\$Script:OfficeVer\OneNote\OpenNotebooks" } })
        $steps.Add(@{ Name='Verify OneNote notebooks';  Action={ Test-OneNoteRestore } })
        $steps.Add(@{ Name='Wait for Outlook profile';  Action={ Wait-ForOutlookProfile }; Verify={ Test-Path $Script:Reg.OutlookPS } })
        $steps.Add(@{ Name='Import Outlook profile';    Action={ Import-OutlookProfile -Name 'OldPcOutlook' }; Verify={ Test-Path $Script:Reg.OutlookPS } })

        Invoke-Pipeline -Steps $steps.ToArray() -Title 'RESTORE SUMMARY'
    } finally { Stop-SwapLog }
}

#endregion

# =====================================================================================
#region MENU / ENTRY POINT
# =====================================================================================

function Show-Menu {
    Update-HealthStatus | Out-Null
    $items = [ordered]@{
        '1' = @{ Group = 'Backup';  Text = 'Full Backup';                 Action = { Start-SwapBackup } }
        '2' = @{ Group = 'Restore'; Text = 'Restore (choose a backup)';   Action = { Start-SwapRestore } }
        '3' = @{ Group = 'Tools';   Text = 'Repair Outlook profile';      Action = { Repair-OutlookProfile } }
        '4' = @{ Group = 'Tools';   Text = 'Show installed apps';         Action = { Get-InstalledAppList } }
        '5' = @{ Group = 'Tools';   Text = 'Connectivity / health check'; Action = { Show-HealthDetail } }
    }

    while ($true) {
        Clear-Host
        Show-AppHeader
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
        Write-Host '    R. Refresh status        Q. Quit' -ForegroundColor DarkGray
        Write-Host ''

        $sel = (Read-Host ' Select an option').Trim().ToUpper()
        if ($sel -eq 'Q') { break }
        if ($sel -eq 'R') { Update-HealthStatus | Out-Null; continue }
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
