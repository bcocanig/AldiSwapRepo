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
    '*ntel*'
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
    'Okta Device Access*',
    'Okta Verify*',
    'OktaVerify-x64-*',
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

$apps = foreach ($path in $paths) {
    if (Test-Path $path) {
        Get-ChildItem $path -ErrorAction Stop | ForEach-Object {
            $item = Get-ItemProperty $_.PsPath -ErrorAction Stop
            #Write-Host "Processing: $($_.PsPath)"
            $item = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
            if (-not $item) { return }

            $displayName = $null
            $prop = $item.PSObject.Properties['DisplayName']
            if ($prop) { $displayName = $prop.Value }

            if ([string]::IsNullOrWhiteSpace($displayName)) {
                # Missing/empty DisplayName
                return
            }

            # Now use $displayName (not $item.DisplayName)
            [PSCustomObject]@{
                DisplayName     = $displayName
            }

        }
    }
}

# Remove duplicate DisplayName entries
$appsUnique = $apps | Sort-Object DisplayName -Unique

$pythonCore = $appsUnique |
    Where-Object { $_.DisplayName -like 'Python 3.*Core Interpreter*' } |
    Sort-Object DisplayName -Descending |
    Select-Object -First 1
# Exclude the ones you don’t want to see
$appsFiltered = $appsUnique |
    Where-Object {
        $name = $_.DisplayName
        # Exclude anything that starts with a lowercase letter
        if ($name -and ($name[0] -cmatch '[a-z]')) {
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
    Select-Object DisplayName |
    Format-Table -AutoSize
