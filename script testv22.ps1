#Requires -Version 5.1
<#
.SYNOPSIS
    Laptop Migration Tool - Backup and Restore user data during laptop swaps.

.DESCRIPTION
    Handles backup of: OneNote notebooks, Outlook registry, email signatures,
    Quick Access pins, Downloads, wallpaper, computer info, and app inventory.
    Supports restore via OneDrive or F Drive.

.NOTES
    Author  : Brandon Cocanig / Chris Zeyen / Paul Aguilera
    Updated : 2025 - Full rewrite for stability, naming consistency, and UX
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# REGION: GLOBAL CONFIGURATION
# ============================================================

$Script:WorkingDir   = 'C:\Temp\LaptopTransferBackups'
$Script:LogDir       = Join-Path $Script:WorkingDir 'Logs'
$Script:LogFile      = $null   # Set when log session opens

# OneNote COM interop state (lazy-initialized)
$Script:OneNoteApp   = $null
$Script:OneNoteXmlNs = $null
$Script:OneNoteSchema = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

$Script:Divider = '=' * 60

# ============================================================
# REGION: LOGGING
# ============================================================

function Write-Log {
    <#
    .SYNOPSIS Writes a timestamped message to both console and the session log file.
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO',
        [System.ConsoleColor]$Color
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "[$timestamp][$Level] $Message"

    # Pick color based on level if not overridden
    if (-not $PSBoundParameters.ContainsKey('Color')) {
        $Color = switch ($Level) {
            'WARN'    { [System.ConsoleColor]::Yellow }
            'ERROR'   { [System.ConsoleColor]::Red    }
            'SUCCESS' { [System.ConsoleColor]::Green  }
            default   { [System.ConsoleColor]::Cyan   }
        }
    }

    Write-Host $entry -ForegroundColor $Color

    if ($Script:LogFile) {
        try { Add-Content -Path $Script:LogFile -Value $entry -Encoding UTF8 }
        catch { <# Don't let a log write kill the session #> }
    }
}

function Start-LogSession {
    <#
    .SYNOPSIS Opens a new timestamped log file for the current operation.
    #>
    param([string]$Operation = 'Session')

    if (-not (Test-Path $Script:LogDir)) {
        New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Script:LogFile = Join-Path $Script:LogDir "Log_${Operation}_${stamp}.txt"
    New-Item -Path $Script:LogFile -ItemType File -Force | Out-Null
    Write-Log "Log session started: $($Script:LogFile)"
}

function Stop-LogSession {
    Write-Log "Log session ended: $($Script:LogFile)" -Level SUCCESS
    $Script:LogFile = $null
}

# ============================================================
# REGION: PATH DETECTION
# ============================================================

function Get-OneDrivePath {
    <#
    .SYNOPSIS
        Detects the user's OneDrive root folder.
        Handles standard ALDI-HOFER, ALDI-499 domain suffix, and ALDI DX variants.
    .OUTPUTS [string] or $null
    #>
    $candidates = @(
        "C:\Users\$env:USERNAME\OneDrive - ALDI-HOFER"
        "C:\Users\$env:USERNAME.ALDI-499\OneDrive - ALDI DX"
        "C:\Users\$env:USERNAME.ALDI-499\OneDrive - ALDI-HOFER"
    )

    foreach ($path in $candidates) {
        if (Test-Path $path) {
            Write-Log "OneDrive detected: $path"
            return $path
        }
    }

    Write-Log 'OneDrive folder not found for any known path variant.' -Level WARN
    return $null
}

function Get-FDrivePath {
    <#
    .SYNOPSIS Returns the F Drive backup share path if reachable, else $null.
    #>
    $path = 'F:\usrnew\For IT Support\Laptop Swap Script\Backups'
    if (Test-Path $path) {
        Write-Log "F Drive detected: $path"
        return $path
    }
    Write-Log 'F Drive share not reachable or not present.' -Level WARN
    return $null
}

function Assert-WorkingDirectory {
    <#
    .SYNOPSIS Ensures the working temp directory exists. Creates it if absent.
    #>
    if (-not (Test-Path $Script:WorkingDir)) {
        New-Item -ItemType Directory -Path $Script:WorkingDir -Force | Out-Null
        Write-Log "Created working directory: $($Script:WorkingDir)"
    }
}

function Assert-SubDirectory {
    <#
    .SYNOPSIS Creates a subdirectory under WorkingDir and returns the full path.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $full = Join-Path $Script:WorkingDir $Name
    if (-not (Test-Path $full)) {
        New-Item -ItemType Directory -Path $full -Force | Out-Null
    }
    return $full
}

# ============================================================
# REGION: ONENOTE INTEROP
# ============================================================

function Initialize-OneNoteApp {
    <#
    .SYNOPSIS Loads the OneNote COM interop assembly and creates the application object.
              Idempotent - safe to call multiple times.
    #>
    if ($Script:OneNoteApp) { return }

    try {
        $interop = Get-Item "$env:WinDir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15*\*" -ErrorAction Stop
        Add-Type -LiteralPath $interop.FullName
        $Script:OneNoteApp = New-Object Microsoft.Office.Interop.OneNote.ApplicationClass

        $xmlDoc = New-Object System.Xml.XmlDocument
        $Script:OneNoteXmlNs = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $Script:OneNoteXmlNs.AddNamespace('one', $Script:OneNoteSchema)

        $ver = (Get-Process onenote -ErrorAction SilentlyContinue).ProductVersion
        Write-Log "OneNote COM initialized. Version: $ver"
    }
    catch [System.Runtime.InteropServices.COMException] {
        throw 'Unable to create OneNote COM object. Is OneNote installed and running?'
    }
}

function Get-OneNoteNotebookList {
    <#
    .SYNOPSIS
        Returns an array of [PSCustomObject]@{Name; Path} for all open notebooks.
    #>
    Initialize-OneNoteApp

    [string]$xml = ''
    $xmlDoc = New-Object System.Xml.XmlDocument
    $Script:OneNoteApp.GetHierarchy($null,
        [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,
        [ref]$xml)
    $xmlDoc.LoadXml($xml)

    $nodes = $xmlDoc.SelectNodes('//one:Notebook', $Script:OneNoteXmlNs)
    $results = foreach ($node in $nodes) {
        [PSCustomObject]@{
            Name = $node.GetAttribute('name')
            Path = $node.GetAttribute('path')
        }
    }
    return $results
}

function Save-OneNoteNotebookList {
    <#
    .SYNOPSIS Saves the current notebook list to OneNoteBooks.json in WorkingDir.
    #>
    param([string]$FileName = 'OneNoteBooks.json')

    Assert-WorkingDirectory
    $outputPath = Join-Path $Script:WorkingDir $FileName

    $notebooks = Get-OneNoteNotebookList

    if (-not $notebooks) {
        Write-Log 'No OneNote notebooks found.' -Level WARN
        return
    }

    $notebooks | Format-Table | Out-String | Write-Host
    $notebooks | ConvertTo-Json | Set-Content -Path $outputPath -Encoding UTF8
    Write-Log "Saved $($notebooks.Count) notebook(s) to: $outputPath" -Level SUCCESS
}

function Export-OneNoteRegistry {
    <#
    .SYNOPSIS Exports the OneNote OpenNotebooks registry key to a .reg file.
    #>
    $regKey  = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\OneNote\OpenNotebooks'
    $outDir  = Assert-SubDirectory 'OneNoteReg'
    $outFile = Join-Path $outDir 'OneNoteNotebooks.reg'

    Write-Log "Exporting OneNote registry: $regKey"
    reg export $regKey $outFile /y 2>&1 | Out-Null

    if (Test-Path $outFile) {
        Write-Log "OneNote registry exported to: $outFile" -Level SUCCESS
    } else {
        Write-Log 'OneNote registry export failed.' -Level ERROR
    }
}

function Import-OneNoteRegistry {
    <#
    .SYNOPSIS Stops OneNote, imports the saved registry file, validates result.
    #>
    $regFile = Join-Path $Script:WorkingDir 'OneNoteReg\OneNoteNotebooks.reg'

    if (-not (Test-Path $regFile)) {
        Write-Log "OneNote registry file not found: $regFile" -Level WARN
        return
    }

    Stop-ProcessIfRunning 'OneNote'
    Stop-ProcessIfRunning 'ONENOTEM'
    Wait-ForProcessExit 'OneNote' -TimeoutSeconds 15

    reg import $regFile | Out-Null

    Write-Log "OneNote registry imported from: $regFile" -Level SUCCESS
    
}
    

function New-OneNoteShortcuts {
    <#
    .SYNOPSIS
        Reads OneNoteBooks.json and creates .url (cloud) or .lnk (local) shortcuts
        in a OneNoteBooks_Shortcuts subfolder, then opens the folder for the tech.
    #>
    $sourceFile    = Join-Path $Script:WorkingDir 'OneNoteBooks.json'
    $shortcutsDir  = Join-Path $Script:WorkingDir 'OneNoteBooks_Shortcuts'

    if (-not (Test-Path $sourceFile)) {
        Write-Log "Notebook JSON not found: $sourceFile" -Level WARN
        return
    }

    if (-not (Test-Path $shortcutsDir)) {
        New-Item -ItemType Directory -Path $shortcutsDir | Out-Null
    }

    $notebooks = Get-Content $sourceFile -Raw | ConvertFrom-Json
    $created   = 0

    foreach ($nb in $notebooks) {
        $path = $nb.Path
        $name = $nb.Name

        if ($path -match '^https?://') {
            # Cloud notebook → .url shortcut with onenote: prefix
            $target       = "onenote:$path"
            $shortcutFile = Join-Path $shortcutsDir "$name.url"
            $shell        = New-Object -ComObject WScript.Shell
            $sc           = $shell.CreateShortcut($shortcutFile)
            $sc.TargetPath = $target
            $sc.Save()
            Write-Log "Created URL shortcut: $name"
            $created++
        } elseif (Test-Path $path) {
            # Local notebook → find .onetoc2 and create .lnk
            $toc = Get-ChildItem $path -Recurse -Filter 'Open Notebook.onetoc2' `
                       -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($toc) {
                $folderName   = Split-Path -Leaf (Split-Path $toc.FullName)
                $shortcutFile = Join-Path $shortcutsDir "$folderName - Open Notebook.lnk"
                $shell        = New-Object -ComObject WScript.Shell
                $sc           = $shell.CreateShortcut($shortcutFile)
                $sc.TargetPath = $toc.FullName
                $sc.Save()
                Write-Log "Created LNK shortcut: $folderName"
                $created++
            } else {
                Write-Log "Could not find 'Open Notebook.onetoc2' in: $path" -Level WARN
            }
        } else {
            Write-Log "Notebook path not accessible: $path" -Level WARN
        }
    }

    Write-Log "Shortcuts created: $created of $($notebooks.Count)" -Level SUCCESS
    Invoke-Item $shortcutsDir
}

function Compare-OneNoteNotebooks {
    <#
    .SYNOPSIS
        Opens OneNote, waits for it to load, re-reads notebook list, then diffs
        against the backup JSON to surface any missing notebooks.
    #>
    $backupFile  = Join-Path $Script:WorkingDir 'OneNoteBooks.json'
    $compareFile = Join-Path $Script:WorkingDir 'OneNoteBooksCompare.json'

    if (-not (Test-Path $backupFile)) {
        Write-Log "Backup notebook list not found: $backupFile" -Level WARN
        return
    }

    Write-Log 'Launching OneNote for notebook comparison...'
    Start-Process 'onenote.exe'

    # Poll for OneNote process instead of sleeping blindly
    Wait-ForProcess -Name 'onenote' -TimeoutSeconds 30

    Save-OneNoteNotebookList -FileName 'OneNoteBooksCompare.json'

    $original = (Get-Content $backupFile  -Raw | ConvertFrom-Json) | Select-Object -ExpandProperty Name
    $current  = (Get-Content $compareFile -Raw | ConvertFrom-Json) | Select-Object -ExpandProperty Name

    $missingList = @(
        Compare-Object -ReferenceObject $original -DifferenceObject $current |
        Where-Object { $_.SideIndicator -eq '<=' } |
        Select-Object -ExpandProperty InputObject
    )

    if ($missingList.Count -gt 0) {
        Write-Log "MISSING notebooks detected ($($missingList.Count)):" -Level WARN
        $missingList | ForEach-Object { Write-Log "  - $_" -Level WARN }
    } else {
        Write-Log 'All notebooks accounted for.' -Level SUCCESS
    }
}

# ============================================================
# REGION: OUTLOOK
# ============================================================

function Export-OutlookRegistry {
    <#
    .SYNOPSIS Exports Outlook Profiles registry key to a named .reg file.
    .PARAMETER FileName Name for the .reg file (without extension).
    #>
    param([Parameter(Mandatory)][string]$FileName)

    $regKey  = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles'
    $outDir  = Assert-SubDirectory 'OutlookReg'
    $outFile = Join-Path $outDir "$FileName.reg"

    Write-Log "Exporting Outlook registry to: $outFile"
    reg export $regKey $outFile /y 2>&1 | Out-Null

    if (Test-Path $outFile) {
        Write-Log "Outlook registry exported: $outFile" -Level SUCCESS
    } else {
        Write-Log 'Outlook registry export failed.' -Level ERROR
    }
}

function Import-OutlookRegistry {
    <#
    .SYNOPSIS
        Waits for Outlook to create its profile registry key (proving it's been opened),
        backs up the new PC's Outlook registry, then imports the old PC's registry.
    #>
    $oldRegFile      = Join-Path $Script:WorkingDir 'OutlookReg\OldPcOutlook.reg'
    $outlookRegPath  = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'

    if (-not (Test-Path $oldRegFile)) {
        Write-Log "Old PC Outlook registry not found: $oldRegFile" -Level WARN
        return
    }

    Write-Log 'Waiting for Outlook to initialize its profile registry key...'
    $found = Wait-ForRegistryKey -KeyPath $outlookRegPath -TimeoutSeconds 180

    if (-not $found) {
        Write-Log 'Outlook profile key not detected within timeout. Continuing anyway.' -Level WARN
    }

    # Back up new PC's Outlook config before overwriting
    Export-OutlookRegistry -FileName 'NewPcBackup'

    Stop-ProcessIfRunning 'OUTLOOK'
    Wait-ForProcessExit 'OUTLOOK' -TimeoutSeconds 20

    reg import $oldRegFile | Out-Null
    Write-Log 'Old PC Outlook registry imported.' -Level SUCCESS
}

function Repair-OutlookRegistry {
    <#
    .SYNOPSIS
        Restores the new PC's Outlook registry backup (rollback after a bad import).
        This is the "Fix Outlook" menu option.
    #>
    $regFile = Join-Path $Script:WorkingDir 'OutlookReg\NewPcBackup.reg'

    if (-not (Test-Path $regFile)) {
        Write-Log "New PC Outlook backup not found: $regFile" -Level ERROR
        Write-Log 'Run a Restore first to generate this file.' -Level WARN
        return
    }

    Stop-ProcessIfRunning 'OUTLOOK'
    Wait-ForProcessExit 'OUTLOOK' -TimeoutSeconds 20

    reg import $regFile 2>&1 | Out-Null
    Write-Log 'New PC Outlook registry restored from backup.' -Level SUCCESS
}

# ============================================================
# REGION: EMAIL SIGNATURES
# ============================================================

function Backup-EmailSignatures {
    $src  = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Signatures"
    $dest = Assert-SubDirectory 'EmailSignatures'

    if (-not (Test-Path $src)) {
        Write-Log "Signatures folder not found: $src" -Level WARN
        return
    }

    robocopy $src $dest /E /COPY:DAT /NJH /NJS /NP | Out-Null
    Write-Log 'Email signatures backed up.' -Level SUCCESS
}

function Restore-EmailSignatures {
    $src  = Join-Path $Script:WorkingDir 'EmailSignatures'
    $dest = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Signatures"

    if (-not (Test-Path $src)) {
        Write-Log "Email signatures backup not found: $src" -Level WARN
        return
    }

    if (-not (Test-Path $dest)) {
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
    }

    robocopy $src $dest /E /COPY:DAT /NJH /NJS /NP | Out-Null
    Write-Log 'Email signatures restored.' -Level SUCCESS
}

# ============================================================
# REGION: QUICK ACCESS
# ============================================================

function Backup-QuickAccess {
    $src  = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"
    $dest = Assert-SubDirectory 'QuickAccessBK'

    if (-not (Test-Path $src)) {
        Write-Log 'Quick Access source folder not found.' -Level WARN
        return
    }
    
    $files = @(Get-ChildItem $src)
    $total = $files.Count
    $i     = 0

    foreach ($file in $files) {
        $i++
        $pct = [int](($i / $total) * 100)
        Write-Progress -Activity 'Backing up Quick Access' `
                       -Status "$($file.Name)" `
                       -PercentComplete $pct

        Copy-Item $file.FullName (Join-Path $dest $file.Name) -Force
    }
    Write-Progress -Activity 'Backing up Quick Access' -Completed
    Write-Log "Quick Access backed up ($total files)." -Level SUCCESS
}

function Restore-QuickAccess {
    $src  = Join-Path $Script:WorkingDir 'QuickAccessBK'
    $dest = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"

    if (-not (Test-Path $src)) {
        Write-Log 'Quick Access backup folder not found. Skipping.' -Level WARN
        return
    }

    robocopy $src $dest /IS /NJH /NJS /NP | Out-Null

    # Restart Explorer to apply pinned items
    Write-Log 'Restarting Explorer to apply Quick Access pins...'
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Wait-ForProcess -Name 'explorer' -TimeoutSeconds 15
    Write-Log 'Quick Access restored.' -Level SUCCESS
}

# ============================================================
# REGION: DOWNLOADS BACKUP
# ============================================================

function Backup-Downloads {
    $src  = "$env:USERPROFILE\Downloads"
    $dest = Assert-SubDirectory 'DownloadFiles'

    $totalBytes = (Get-ChildItem $src -Recurse -ErrorAction SilentlyContinue |
                   Measure-Object -Property Length -Sum).Sum
    $totalGB    = [math]::Round($totalBytes / 1GB, 2)

    if ($totalGB -gt 1) {
        $answer = Read-Host "Downloads folder is ${totalGB}GB. Back it up? [Y/N]"
        if ($answer.Trim().ToUpper() -ne 'Y') {
            Write-Log 'Downloads backup skipped by user.'
            return
        }
    }

    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $zipFile = Join-Path $dest "Downloads_${stamp}.zip"
    try {
        Write-Log "Compressing Downloads to: $zipFile"
        Compress-Archive -Path "$src\*" -DestinationPath $zipFile -Force `
                        -CompressionLevel Fastest
        Write-Log 'Downloads backup complete.' -Level SUCCESS
    }
    catch {
        Write-Log "Downloads backup failed: $_" -Level ERROR
    }
}

# ============================================================
# REGION: WALLPAPER
# ============================================================

function Backup-Wallpaper {
    $regPath   = 'HKCU:\Control Panel\Desktop'
    $wallpaper = (Get-ItemProperty $regPath -Name Wallpaper -ErrorAction SilentlyContinue).Wallpaper
    $dest      = Assert-SubDirectory 'Wallpaper'

    if ($wallpaper -and (Test-Path $wallpaper)) {
        Copy-Item $wallpaper $dest -Force
        Write-Log "Wallpaper backed up: $wallpaper" -Level SUCCESS
    } else {
        Write-Log 'No wallpaper found or path invalid.' -Level WARN
    }
}

# ============================================================
# REGION: COMPUTER INFO
# ============================================================

function Save-ComputerInfo {
    Assert-WorkingDirectory

    $info = [ordered]@{
        'DeviceName'     = $env:COMPUTERNAME
        'OS'             = (Get-CimInstance Win32_OperatingSystem).Caption
        'RAM_GB'         = "{0:N2}" -f ((Get-CimInstance Win32_PhysicalMemory |
                               Measure-Object Capacity -Sum).Sum / 1GB)
        'Storage'        = (Get-CimInstance Win32_LogicalDisk |
                            Where-Object { $_.DriveType -eq 3 } |
                            ForEach-Object {
                                "$($_.DeviceID): $([math]::Round($_.FreeSpace/1GB,2))GB free of $([math]::Round($_.Size/1GB,2))GB"
                            }) -join '; '
        'Model'          = (Get-CimInstance Win32_ComputerSystem).Model
        'BIOSVersion'    = (Get-CimInstance Win32_BIOS).SMBIOSBIOSVersion
        'ServiceTag'     = (Get-CimInstance Win32_BIOS).SerialNumber
        'CapturedAt'     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        'CapturedByUser' = $env:USERNAME
    }

    $outFile = Join-Path $Script:WorkingDir 'ComputerInfo.json'
    $info | ConvertTo-Json | Set-Content $outFile -Encoding UTF8
    Write-Log "Computer info saved to: $outFile" -Level SUCCESS

    $info.GetEnumerator() | ForEach-Object {
        Write-Log "  $($_.Key): $($_.Value)"
    }
}

# ============================================================
# REGION: FILE TREE
# ============================================================

function Save-FileTree {
    param(
        [string]$Directory  = "$env:USERPROFILE\Downloads",
        [string]$TargetName = 'Downloads'
    )

    $treeDir  = Assert-SubDirectory 'Trees'
    $outFile  = Join-Path $treeDir "$TargetName.txt"

    tree $Directory /f | Out-File $outFile -Encoding UTF8
    Write-Log "File tree saved: $outFile"
}

# ============================================================
# REGION: APP INVENTORY
# ============================================================

function Get-InstalledAppInventory {
    <#
    .SYNOPSIS Checks for known specialty apps and reports which are installed.
    #>

    # Registry paths (without the InstallDate value at the end)
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    # Add per-user uninstall keys from all loaded user hives (HKU)
    $paths += Get-ChildItem 'HKU:\' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match 'S-1-5-21-\d+-\d+-\d+-\d+$$' } |  # filter SIDs
        ForEach-Object {
            "Registry::${($_.Name)}\Software\Microsoft\Windows\CurrentVersion\Uninstall"
        }

    # Names you do NOT want to see (adjust as needed)
    $excludeNames = @(
        '2013.12.13 Enterprise Client Controls*',
        '64 Bit HP CIO Components Installer*',
        '*aldi-support.beyondtrustcloud.com*',
        '*OneDrive*',
        'Active Directory Rights*',
        'Adobe Acrobat Reader*',
        'ALDI Encryption Add-in for Outlook*',
        'ALDI Font SUED OT 1.0.0.0*',
        'ALDI LAN Desk Compliance Status Check Client*',
        'ALDI LANDESK Compliance Check*',
        'ALDI SUED Fonts*',
        'AppProtection*',
        'Assima Application Listener*',
        'Cherry SmartCard Package V3.3 Build 9*',
        'Cirrus Audio*',
        'Citrix*',
        'ClickShare Extension Pack*',
        'CryptoPro*',
        'Customer Support*',
        'Dell*',
        'DFUDriverSetupX64Setup*',
        'DisplayLink Graphics*',
        'Dynamic Application*',
        'Eclipse Temurin JRE with Hotspot*',
        'EU Waste Recycling Information*',
        'Forticlient*',
        'Greenshot 1.3.315*',
        'Information Center*',
        'Ivanti*',
        'Jabra*',
        'LANDESK Advance Agent*',
        'Microsoft .NET*',
        'Microsoft 365*',
        'Microsoft ASP*',
        'Microsoft Device*',
        'Microsoft Edge*',
        'Microsoft Intune*',
        'Microsoft Purview Information Protection*',
        'Microsoft SQL Server 2008*',
        'Microsoft SQL Server 2012*',
        'Microsoft Teams*',
        'Microsoft Visio Viewer 2016*',
        'Microsoft Visual C++*',
        'Microsoft Visual Studio 2010 Tools for Office Runtime (x64)*',
        'Microsoft Windows*',
        'MTOP Client*',
        'Nagyv*llalati*',
        'Office 16 Click-to-Run Extensibility Component*',
        'Office 16 Click-to-Run Localization Component*',
        'Okta Device Access 6.4.0.0*',
        'Okta Verify*',
        'OktaVerify-x64-6.4.0.0*',
        'Online Plug-in*',
        'Phish Alert*',
        'PowerToys (Preview) x64*',
        'Programi Microsoft 365 za podjetja - sl-si*',
        'Programi Microsoft 365 za podjetja - sl-si.proof*',
        'Realtek*',
        'Required Runtimes*',
        'SAP Crystal Reports runtime engine for .NET Framework (32-bit)*',
        'SAP Crystal Reports runtime engine for .NET Framework (64-bit)*',
        'Self-service Plug-in*',
        'Silverfort Client*',
        'Skyhigh Client Proxy*',
        'SQL Server-Berichts-Generator 3 für SQL Server 2014*',
        'SQL Server-Berichts-Generator*',
        'TbtLegacy*',
        'Teams Machine-Wide Installer*',
        'Thunderbolt*',
        'Trellix Agent*',
        'Trellix Data Exchange Layer for TA*',
        'Trend Micro Apex One Security Agent*',
        'USB Drive Letter Manager (x64)*',
        'Update for*',
        'Visual Studio Tools for the Office system 3.0 Runtime*'
        # add more patterns or exact names
    )
    # Grabs properties for each reg entry found
    $apps = foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
                $item = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
                if ($item.DisplayName) {
                    [PSCustomObject]@{
                        DisplayName     = $item.DisplayName
                        InstallLocation = $item.InstallLocation
                        Publisher       = $item.Publisher
                        InstallDate     = $item.InstallDate
                        Version         = $item.DisplayVersion
                        InstallSource   = $item.ModifyPath
                    }
                        #RegistryKey = $_.PsPath
                }
            }
        }
    }

    # Remove duplicate DisplayName entries
    $appsUnique = $apps | Sort-Object DisplayName -Unique
    # Filter out other Python entries
    $pythonCore = $appsUnique |
        Where-Object { $_.DisplayName -like 'Python 3.*Core Interpreter*' } |
        Sort-Object DisplayName -Descending |
        Select-Object -First 1
    # Exclude the ones you don’t want to see
    $appsFiltered = $appsUnique |
        Where-Object {
            $name = $_.DisplayName
            $publisher = $_.Publisher
            # Exclude anything that starts with a lowercase letter
            if ($name -and ($name[0] -cmatch '[a-z]')) {
                return $false
            }
            if ($publisher -like '*Intel*') {
                return $false
            }
            # Drop all Python 3.* except Core Interpreter
            if ($name -like 'Python 3.*') {
                $_ -eq $pythonCore
            }
            else{
                -not ($excludeNames | Where-Object { $name -like $_ })
            }
        }

    # Show result (adjust properties/format as needed)
    $appsFiltered |
        Sort-Object DisplayName |
        Select-Object DisplayName, Publisher |
        Format-Table -AutoSize
}

# ============================================================
# REGION: BACKUP DESTINATION - ONEDRIVE / F DRIVE
# ============================================================

function Backup-ToOneDrive {
    $oneDrive = Get-OneDrivePath
    if (-not $oneDrive) {
        Write-Log 'Cannot back up to OneDrive - path not found.' -Level ERROR
        return
    }

    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dest   = Join-Path $oneDrive "Backup_$stamp"
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    Write-Log "Copying backup to OneDrive: $dest"
    robocopy $Script:WorkingDir $dest /E /NJH /NJS /NP | Out-Null
    Write-Log "OneDrive backup complete: $dest" -Level SUCCESS
}

function Backup-ToFDrive {
    $fDrive = Get-FDrivePath
    if (-not $fDrive) {
        Write-Log 'F Drive not available - skipping F Drive backup.' -Level WARN
        return
    }

    $dest = Join-Path $fDrive "Backup_$env:USERNAME"
    if (Test-Path $dest) {
        # Remove stale username backup before overwriting
        Remove-Item $dest -Recurse -Force
    }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    Write-Log "Copying backup to F Drive: $dest"
    robocopy $Script:WorkingDir $dest /E /NJH /NJS /NP /XD "DownloadFiles" "Wallpaper" "Trees"| Out-Null
    Write-Log 'F Drive backup complete.' -Level SUCCESS
}

# ============================================================
# REGION: RESTORE SOURCE SELECTION
# ============================================================

function Select-OneDriveBackup {
    <#
    .SYNOPSIS
        Lists timestamped backup folders in OneDrive, lets the tech pick one,
        and copies its contents into the working directory.
    #>
    $oneDrive = Get-OneDrivePath
    if (-not $oneDrive) {
        throw 'OneDrive not found. Cannot proceed with OneDrive restore.'
    }

    $folders = @(Get-ChildItem $oneDrive -Directory |
                 Where-Object { $_.Name -match '^Backup_\d{8}_\d{6}$' } |
                 Sort-Object LastWriteTime -Descending)

    if ($folders.Count -eq 0) {
        throw "No Backup_yyyyMMdd_HHmmss folders found in: $oneDrive"
    }

    Write-Host ''
    Write-Host 'Available OneDrive backups:'
    for ($i = 0; $i -lt $folders.Count; $i++) {
        $ts = $folders[$i].Name -replace 'Backup_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})',
                                          '$1-$2-$3 $4:$5:$6'
        Write-Host "  [$i] $ts"
    }

    $raw = Read-Host 'Select backup number (Enter = 0 for latest)'
    $idx = if ($raw -match '^\d+$') { [int]$raw } else { 0 }
    if ($idx -ge $folders.Count) { $idx = 0 }

    $chosen = $folders[$idx].FullName

    # Support both flat and nested (LaptopTransferBackups subfolder) layout
    $nested = Join-Path $chosen 'LaptopTransferBackups'
    $source = if (Test-Path $nested) { $nested } else { $chosen }

    Write-Log "Restoring from: $source"
    Assert-WorkingDirectory

    robocopy $source $Script:WorkingDir /E /NJH /NJS /NP | Out-Null
    Write-Log 'OneDrive backup contents copied to working directory.' -Level SUCCESS

    return $source
}

function Select-FDriveBackup {
    <#
    .SYNOPSIS Copies the user's named backup from the F Drive to the working directory.
    #>
    $fDrive = Get-FDrivePath
    if (-not $fDrive) {
        throw 'F Drive not found. Cannot proceed with F Drive restore.'
    }

    $source = Join-Path $fDrive "Backup_$env:USERNAME"
    if (-not (Test-Path $source)) {
        throw "F Drive backup not found for user '$env:USERNAME': $source"
    }

    Write-Log "Restoring from F Drive: $source"
    Assert-WorkingDirectory

    robocopy $source $Script:WorkingDir /E /NJH /NJS /NP | Out-Null

    # Clean up F Drive source after successful copy
    Remove-Item $source -Recurse -Force
    Write-Log 'F Drive backup copied and source removed.' -Level SUCCESS
}

# ============================================================
# REGION: PROCESS HELPERS
# ============================================================

function Stop-ProcessIfRunning {
    param([Parameter(Mandatory)][string]$Name)

    $proc = Get-Process -Name $Name -ErrorAction SilentlyContinue
    if ($proc) {
        Write-Log "Stopping process: $Name"
        Stop-Process -Name $Name -Force -ErrorAction SilentlyContinue
    }
}

function Wait-ForProcessExit {
    <#
    .SYNOPSIS Polls until the named process is no longer running or timeout expires.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = 30
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (-not (Get-Process -Name $Name -ErrorAction SilentlyContinue)) {
            Write-Log "$Name has exited."
            return $true
        }
        Start-Sleep -Milliseconds 500
    }
    Write-Log "Timeout waiting for $Name to exit." -Level WARN
    return $false
}

function Wait-ForProcess {
    <#
    .SYNOPSIS Polls until the named process is running or timeout expires.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$TimeoutSeconds = 30
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (Get-Process -Name $Name -ErrorAction SilentlyContinue) {
            Write-Log "$Name is running."
            return $true
        }
        Start-Sleep -Milliseconds 500
    }
    Write-Log "Timeout waiting for $Name to start." -Level WARN
    return $false
}

function Wait-ForRegistryKey {
    <#
    .SYNOPSIS Polls until the given HKCU/HKLM registry key exists or timeout expires.
    .OUTPUTS [bool]
    #>
    param(
        [Parameter(Mandatory)][string]$KeyPath,
        [int]$TimeoutSeconds = 120
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (Test-Path $KeyPath) {
            Write-Log "Registry key found: $KeyPath"
            return $true
        }
        Start-Sleep -Milliseconds 1000
    }
    Write-Log "Registry key not found within ${TimeoutSeconds}s: $KeyPath" -Level WARN
    return $false
}

# ============================================================
# REGION: PRE-BACKUP CLEANUP CHECK
# ============================================================

function Invoke-WorkingDirectoryCheck {
    <#
    .SYNOPSIS
        If the working directory already exists, prompts the tech to Delete, Skip, or Quit.
        Prevents accidental merge of old and new backup data.
    #>
    if (-not (Test-Path $Script:WorkingDir)) { return }

    Write-Log "Working directory already exists: $($Script:WorkingDir)" -Level WARN
    Write-Log 'It is recommended to remove it before starting a fresh backup.' -Level WARN

    do {
        $choice = (Read-Host '[D]elete, [S]kip, or [Q]uit?').Trim().ToUpper()
        switch ($choice) {
            'D' {
                Remove-Item $Script:WorkingDir -Recurse -Force
                Write-Log 'Working directory deleted.' -Level SUCCESS
            }
            'S' { Write-Log 'Skipping removal. Existing data may be overwritten.' -Level WARN }
            'Q' { Write-Log 'Exiting at user request.'; exit 0 }
            default { Write-Host 'Enter D, S, or Q.' }
        }
    } while ($choice -notin 'D','S','Q')
}

# ============================================================
# REGION: CONNECTIVITY CHECKS
# ============================================================

function Assert-OneDriveConnected {
    $path = Get-OneDrivePath
    if (-not $path) {
        Write-Log 'OneDrive is not signed in or syncing.' -Level WARN
        Write-Host ''
        Write-Host 'Please sign in to OneDrive:'
        Write-Host '  1. Open: C:\Users\<username>\AppData\Local\Microsoft\OneDrive\OneDrive.exe'
        Write-Host '  2. Sign in with your work account'
        Write-Host '  3. Wait for initial sync to complete'
        Write-Host ''
        $continue = Read-Host 'Press Enter when OneDrive is ready, or type Q to quit'
        if ($continue.Trim().ToUpper() -eq 'Q') { exit 0 }

        $path = Get-OneDrivePath
        if (-not $path) {
            throw 'OneDrive still not detected. Please resolve and re-run.'
        }
    }
    return $path
}

function Assert-FDriveConnected {
    $path = Get-FDrivePath
    if (-not $path) {
        Write-Log 'F Drive backup share not found.' -Level WARN
    }
    return $path
}

# ============================================================
# REGION: ORCHESTRATORS
# ============================================================

function Start-Backup {
    $TotalSteps = 13
    Write-Host $Script:Divider
    Write-Host '  FULL AUTO BACKUP'
    Write-Host $Script:Divider

    Start-LogSession -Operation 'Backup'
    Write-Log 'Backup started.'

    try {
        Write-Log "--- Step 1/${TotalSteps}: Checking working directory"
        Invoke-WorkingDirectoryCheck
        Start-LogSession -Operation 'Backup'
        Write-Log 'Directory Deleted, New Log Started.'
        Assert-WorkingDirectory

        Write-Log "--- Step 2/${TotalSteps}: Checking OneDrive"
        Assert-OneDriveConnected | Out-Null

        Write-Log "--- Step 3/${TotalSteps}: Checking F Drive"
        Assert-FDriveConnected | Out-Null

        Write-Log "--- Step 4/${TotalSteps}: Saving computer info"
        Save-ComputerInfo

        Write-Log "--- Step 5/${TotalSteps}: Backing up email signatures"
        Backup-EmailSignatures

        Write-Log "--- Step 6/${TotalSteps}: Capturing OneNote notebooks"
        Save-OneNoteNotebookList

        Write-Log "--- Step 7/${TotalSteps}: Creating OneNote shortcuts"
        New-OneNoteShortcuts

        Write-Log "--- Step 8/${TotalSteps}: Exporting Outlook registry"
        Export-OutlookRegistry -FileName "OldPcOutlook"

        Write-Log "--- Step 9/${TotalSteps}: Exporting OneNote registry"
        Export-OneNoteRegistry

        Write-Log "--- Step 10/${TotalSteps}: Backing up Quick Access"
        Backup-QuickAccess

        Write-Log "--- Step 11/${TotalSteps}: Backing up Downloads"
        Backup-Downloads

        Write-Log "--- Step 12/${TotalSteps}: Backing up wallpaper"
        Backup-Wallpaper

        Write-Log "--- Step 13/${TotalSteps}: Saving file tree"
        Save-FileTree

        Write-Log '--- Copying to OneDrive'
        Backup-ToOneDrive

        Write-Log '--- Copying to F Drive (if available)'
        Backup-ToFDrive

        Write-Log '--- Generating app inventory'
        Get-InstalledAppInventory

        Write-Log 'BACKUP COMPLETE.' -Level SUCCESS
        Write-Log "NOTE: $($Script:WorkingDir) has NOT been deleted automatically." -Level WARN
        Write-Log 'Manually delete it when the new laptop is confirmed working.'
    }
    catch {
        Write-Log "BACKUP FAILED: $_" -Level ERROR
    }
    finally {
        Stop-LogSession
    }
}

function Start-OneDriveRestore {
    $TotalSteps = 9
    Write-Host $Script:Divider
    Write-Host '  FULL RESTORE - OneDrive'
    Write-Host $Script:Divider

    Start-LogSession -Operation 'Restore_OneDrive'
    Write-Log 'OneDrive restore started.'

    try {
        Write-Log "--- Step 1/${TotalSteps}: Launching Outlook (starts profile creation)"
        Start-Process 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE' `
                      -ErrorAction SilentlyContinue

        Write-Log "--- Step 2/${TotalSteps}: Verifying OneDrive"
        Assert-OneDriveConnected | Out-Null

        Write-Log "--- Step 3/${TotalSteps}: Selecting backup from OneDrive"
        Select-OneDriveBackup

        Write-Log "--- Step 4/${TotalSteps}: Restoring email signatures"
        Restore-EmailSignatures

        Write-Log "--- Step 5/${TotalSteps}: Restoring Quick Access pins"
        Restore-QuickAccess

        Write-Log "--- Step 6/${TotalSteps}: Importing OneNote registry"
        Import-OneNoteRegistry

        Write-Log "--- Step 7/${TotalSteps}: Comparing OneNote Notebooks"
        Compare-OneNoteNotebooks

        Write-Log "--- Step 8/${TotalSteps}: Importing Outlook registry"
        Import-OutlookRegistry

        Write-Log "--- Step 9/${TotalSteps}: Running data integrity check"
        Invoke-IntegrityCheck

        Write-Log 'RESTORE COMPLETE.' -Level SUCCESS
        Write-Log "Backup files remain at: $($Script:WorkingDir)" -Level WARN
        Write-Log 'Verify the new laptop is working, then delete that folder manually.'
    }
    catch {
        Write-Log "RESTORE FAILED: $_" -Level ERROR
    }
    finally {
        Stop-LogSession
    }
}

function Start-FDriveRestore {
    $TotalSteps = 9
    Write-Host $Script:Divider
    Write-Host '  FULL RESTORE - F Drive'
    Write-Host $Script:Divider

    Start-LogSession -Operation 'Restore_FDrive'
    Write-Log 'F Drive restore started.'

    try {
        Write-Log "--- Step 1/${TotalSteps}: Launching Outlook"
        Start-Process 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE' `
                      -ErrorAction SilentlyContinue

        Write-Log "--- Step 2/${TotalSteps}: Verifying F Drive"
        Assert-FDriveConnected | Out-Null

        Write-Log "--- Step 3/${TotalSteps}: Copying backup from F Drive"
        Select-FDriveBackup

        Write-Log "--- Step 4/${TotalSteps}: Restoring email signatures"
        Restore-EmailSignatures

        Write-Log "--- Step 5/${TotalSteps}: Restoring Quick Access pins"
        Restore-QuickAccess

        Write-Log "--- Step 6/${TotalSteps}: Importing OneNote registry"
        Import-OneNoteRegistry

        Write-Log "--- Step 7/${TotalSteps}: Comparing OneNote Notebooks"
        Compare-OneNoteNotebooks

        Write-Log "--- Step 8/${TotalSteps}: Importing Outlook registry"
        Import-OutlookRegistry

        Write-Log "--- Step 9/${TotalSteps}: Running data integrity check"
        Invoke-IntegrityCheck

        Write-Log 'F DRIVE RESTORE COMPLETE.' -Level SUCCESS
    }
    catch {
        Write-Log "F DRIVE RESTORE FAILED: $_" -Level ERROR
    }
    finally {
        Stop-LogSession
    }
}

function Start-OneDriveRestoreWithoutOutlook {
    <#
    .SYNOPSIS OneDrive restore that skips the Outlook registry step.
    #>
    $TotalSteps = 7
    Write-Host $Script:Divider
    Write-Host '  RESTORE - OneDrive (Outlook-free)'
    Write-Host $Script:Divider

    Start-LogSession -Operation 'Restore_OneDrive_NoOutlook'

    try {
        Write-Log "--- Step 1/${TotalSteps}: Verifying OneDrive"
        Assert-OneDriveConnected | Out-Null

        Write-Log "--- Step 2/${TotalSteps}: Selecting backup"
        Select-OneDriveBackup

        Write-Log "--- Step 3/${TotalSteps}: Restoring email signatures"
        Restore-EmailSignatures

        Write-Log "--- Step 4/${TotalSteps}: Restoring Quick Access"
        Restore-QuickAccess

        Write-Log "--- Step 5/${TotalSteps}: Importing OneNote registry"
        Import-OneNoteRegistry

        Write-Log "--- Step 6/${TotalSteps}: Comparing OneNote Notebooks"
        Compare-OneNoteNotebooks

        Write-Log "--- Step 7/${TotalSteps}: Running data integrity check"
        Invoke-IntegrityCheck -OutlookFree

        Write-Log 'RESTORE (NO OUTLOOK) COMPLETE.' -Level SUCCESS
    }
    catch {
        Write-Log "RESTORE FAILED: $_" -Level ERROR
    }
    finally {
        Stop-LogSession
    }
}

# ============================================================
# REGION: DATA INTEGRITY CHECK
# ============================================================

function Invoke-IntegrityCheck {
    <#
    .SYNOPSIS
        Compares what was backed up against the restored state on the new device.
        Prints a PASS / WARN / FAIL row for each data category and a summary line.
    .PARAMETER OutlookFree
        Pass this switch when running an Outlook-free restore so the Outlook
        profile check is skipped rather than reported as a failure.
    #>
    param([switch]$OutlookFree)

    $checks = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ── 1. Email Signatures ──────────────────────────────────
    $sigBackup = Join-Path $Script:WorkingDir 'EmailSignatures'
    $sigDest   = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Signatures"

    if (Test-Path $sigBackup) {
        $backupCount = @(Get-ChildItem $sigBackup -Recurse -File).Count
        if (Test-Path $sigDest) {
            $destCount = @(Get-ChildItem $sigDest -Recurse -File).Count
            if ($destCount -ge $backupCount) {
                $checks.Add([PSCustomObject]@{ Check = 'Email Signatures';   Status = 'PASS'; Detail = "$backupCount file(s) backed up - $destCount present on device" })
            } else {
                $checks.Add([PSCustomObject]@{ Check = 'Email Signatures';   Status = 'WARN'; Detail = "Backup had $backupCount file(s) - only $destCount found on device" })
            }
        } else {
            $checks.Add([PSCustomObject]@{ Check = 'Email Signatures';       Status = 'FAIL'; Detail = "Destination folder missing: $sigDest" })
        }
    } else {
        $checks.Add([PSCustomObject]@{ Check = 'Email Signatures';           Status = 'WARN'; Detail = 'No backup found - user may not have had any signatures' })
    }

    # ── 2. Quick Access Pins ─────────────────────────────────
    $qaBackup = Join-Path $Script:WorkingDir 'QuickAccessBK'
    $qaDest   = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"

    if (Test-Path $qaBackup) {
        $backupCount = @(Get-ChildItem $qaBackup -File).Count
        if (Test-Path $qaDest) {
            $destCount = @(Get-ChildItem $qaDest -File).Count
            if ($destCount -ge $backupCount) {
                $checks.Add([PSCustomObject]@{ Check = 'Quick Access Pins';  Status = 'PASS'; Detail = "$backupCount file(s) backed up - $destCount present in destination" })
            } else {
                $checks.Add([PSCustomObject]@{ Check = 'Quick Access Pins';  Status = 'WARN'; Detail = "Backup had $backupCount file(s) - only $destCount found in destination" })
            }
        } else {
            $checks.Add([PSCustomObject]@{ Check = 'Quick Access Pins';      Status = 'FAIL'; Detail = "Destination folder missing: $qaDest" })
        }
    } else {
        $checks.Add([PSCustomObject]@{ Check = 'Quick Access Pins';          Status = 'WARN'; Detail = 'No Quick Access backup found' })
    }

    # ── 3. OneNote Notebooks ─────────────────────────────────
    $notebookJson = Join-Path $Script:WorkingDir 'OneNoteBooks.json'

    if (Test-Path $notebookJson) {
        try {
            $backedUp = @((Get-Content $notebookJson -Raw | ConvertFrom-Json) | Select-Object -ExpandProperty Name)
            $current  = @((Get-OneNoteNotebookList) | Select-Object -ExpandProperty Name)
            $missing  = @(
                Compare-Object -ReferenceObject $backedUp -DifferenceObject $current |
                Where-Object   { $_.SideIndicator -eq '<=' } |
                Select-Object  -ExpandProperty InputObject
            )
            if ($missing.Count -eq 0) {
                $checks.Add([PSCustomObject]@{ Check = 'OneNote Notebooks';  Status = 'PASS'; Detail = "All $($backedUp.Count) notebook(s) confirmed open in OneNote" })
            } else {
                $checks.Add([PSCustomObject]@{ Check = 'OneNote Notebooks';  Status = 'WARN'; Detail = "$($missing.Count) missing: $($missing -join ', ')" })
            }
        } catch {
            $checks.Add([PSCustomObject]@{ Check = 'OneNote Notebooks';      Status = 'WARN'; Detail = "Could not verify - OneNote may still be loading ($_)" })
        }
    } else {
        $checks.Add([PSCustomObject]@{ Check = 'OneNote Notebooks';          Status = 'WARN'; Detail = 'OneNoteBooks.json not found in backup' })
    }

    # ── 4. OneNote Registry ──────────────────────────────────
    $oneNoteRegFile = Join-Path $Script:WorkingDir 'OneNoteReg\OneNoteNotebooks.reg'
    $oneNoteRegKey  = 'HKCU:\Software\Microsoft\Office\16.0\OneNote\OpenNotebooks'

    if (Test-Path $oneNoteRegFile) {
        if (Test-Path $oneNoteRegKey) {
            $valueCount = @(
                Get-ItemProperty $oneNoteRegKey |
                Get-Member -MemberType NoteProperty |
                Where-Object { $_.Name -notmatch '^PS' }
            ).Count
            $checks.Add([PSCustomObject]@{ Check = 'OneNote Registry';       Status = 'PASS'; Detail = "Key present - $valueCount notebook entry(s) registered" })
        } else {
            $checks.Add([PSCustomObject]@{ Check = 'OneNote Registry';       Status = 'FAIL'; Detail = 'Registry key absent after import - notebooks may not reopen automatically' })
        }
    } else {
        $checks.Add([PSCustomObject]@{ Check = 'OneNote Registry';           Status = 'WARN'; Detail = 'OneNote registry backup file not found' })
    }

    # ── 5. Outlook Profiles ──────────────────────────────────
    if ($OutlookFree) {
        $checks.Add([PSCustomObject]@{ Check = 'Outlook Profiles';           Status = 'PASS'; Detail = 'Skipped - Outlook-free restore mode' })
    } else {
        $outlookRegFile = Join-Path $Script:WorkingDir 'OutlookReg\OldPcOutlook.reg'
        $outlookRegKey  = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'

        if (Test-Path $outlookRegFile) {
            if (Test-Path $outlookRegKey) {
                $profileCount = @(Get-ChildItem $outlookRegKey -ErrorAction SilentlyContinue).Count
                if ($profileCount -gt 0) {
                    $checks.Add([PSCustomObject]@{ Check = 'Outlook Profiles'; Status = 'PASS'; Detail = "$profileCount profile(s) present in registry" })
                } else {
                    $checks.Add([PSCustomObject]@{ Check = 'Outlook Profiles'; Status = 'WARN'; Detail = 'Profiles key exists but is empty - open Outlook to complete profile creation' })
                }
            } else {
                $checks.Add([PSCustomObject]@{ Check = 'Outlook Profiles';   Status = 'FAIL'; Detail = 'Outlook Profiles registry key absent - import may have failed' })
            }
        } else {
            $checks.Add([PSCustomObject]@{ Check = 'Outlook Profiles';       Status = 'WARN'; Detail = 'Outlook registry backup file not found' })
        }
    }

    # ── 6. Downloads Backup ──────────────────────────────────
    $dlDir = Join-Path $Script:WorkingDir 'DownloadFiles'

    if (Test-Path $dlDir) {
        $zips = @(Get-ChildItem $dlDir -Filter '*.zip' -File)
        if ($zips.Count -gt 0) {
            $sizeMB = [math]::Round(($zips | Measure-Object Length -Sum).Sum / 1MB, 1)
            $checks.Add([PSCustomObject]@{ Check = 'Downloads Backup';       Status = 'PASS'; Detail = "$($zips.Count) zip(s) present - ${sizeMB} MB total. Restore manually from: $dlDir" })
        } else {
            $checks.Add([PSCustomObject]@{ Check = 'Downloads Backup';       Status = 'WARN'; Detail = 'Backup folder exists but contains no zip files' })
        }
    } else {
        $checks.Add([PSCustomObject]@{ Check = 'Downloads Backup';           Status = 'WARN'; Detail = 'Downloads were not backed up or the folder was skipped' })
    }

    # ── Print results ────────────────────────────────────────
    $passCount = @($checks | Where-Object { $_.Status -eq 'PASS' }).Count
    $warnCount = @($checks | Where-Object { $_.Status -eq 'WARN' }).Count
    $failCount = @($checks | Where-Object { $_.Status -eq 'FAIL' }).Count

    Write-Host ''
    Write-Host $Script:Divider
    Write-Host '  DATA INTEGRITY CHECK'
    Write-Host $Script:Divider

    foreach ($c in $checks) {
        $color  = switch ($c.Status) { 'PASS' { 'Green' } 'WARN' { 'Yellow' } default { 'Red' } }
        $symbol = switch ($c.Status) { 'PASS' { '[PASS]' } 'WARN' { '[WARN]' } default { '[FAIL]' } }
        Write-Host ('  {0}  {1,-22}  {2}' -f $symbol, $c.Check, $c.Detail) -ForegroundColor $color
        $logLevel = switch ($c.Status) { 'PASS' { 'SUCCESS' } 'WARN' { 'WARN' } default { 'ERROR' } }
        Write-Log "$symbol $($c.Check): $($c.Detail)" -Level $logLevel
    }

    Write-Host ''
    $summaryColor = if ($failCount -gt 0) { 'Red' } elseif ($warnCount -gt 0) { 'Yellow' } else { 'Green' }
    Write-Host ('  {0} passed    {1} warnings    {2} failed' -f $passCount, $warnCount, $failCount) -ForegroundColor $summaryColor
    Write-Host $Script:Divider
    Write-Host ''
}

# ============================================================
# REGION: MENU
# ============================================================

function Show-Menu {
    Write-Host ''
    Write-Host $Script:Divider
    Write-Host '  LAPTOP MIGRATION TOOL'
    Write-Host $Script:Divider
    Write-Host ''
    Write-Host '  BACKUP'
    Write-Host '  [1]  Full Backup (OneDrive + F Drive)'
    Write-Host ''
    Write-Host '  RESTORE'
    Write-Host '  [2]  Restore from OneDrive'
    Write-Host '  [3]  Restore from F Drive'
    Write-Host '  [4]  Restore from OneDrive (skip Outlook)'
    Write-Host ''
    Write-Host '  UTILITIES'
    Write-Host '  [5]  Fix Outlook Registry (rollback bad import)'
    Write-Host '  [6]  Show Installed App Inventory'
    Write-Host '  [7]  Test OneDrive Connectivity'
    Write-Host '  [8]  Test F Drive Connectivity'
    Write-Host ''
    Write-Host '  [Q]  Quit'
    Write-Host ''
}

# ============================================================
# REGION: ENTRY POINT
# ============================================================

try {
    while ($true) {
        Show-Menu
        $selection = (Read-Host 'Selection').Trim().ToUpper()

        switch ($selection) {
            '1' { Start-Backup }
            '2' { Start-OneDriveRestore }
            '3' { Start-FDriveRestore }
            '4' { Start-OneDriveRestoreWithoutOutlook }
            '5' { Repair-OutlookRegistry }
            '6' { Get-InstalledAppInventory }
            '7' {
                $p = Get-OneDrivePath
                if ($p) { Write-Log "OneDrive OK: $p" -Level SUCCESS }
                else     { Write-Log 'OneDrive NOT found.' -Level WARN }
            }
            '8' {
                $p = Get-FDrivePath
                if ($p) { Write-Log "F Drive OK: $p" -Level SUCCESS }
                else     { Write-Log 'F Drive NOT found.' -Level WARN }
            }
            'Q' { Write-Host 'Goodbye.'; exit 0 }
            default { Write-Host "  Invalid selection: '$selection'" -ForegroundColor Yellow }
        }
    }
} catch {
    Write-Host ''
    Write-Host "FATAL ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
} finally {
    Write-Host ''
    Read-Host 'Press Enter to close'
}