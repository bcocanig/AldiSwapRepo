# Creator: Brandon Cocanig
# DATE: 11/23/23
# Date updated 6/27 !!PLEASE UPDATE WHEN CHANGES ARE MADE!!
# Updated: Chris Zeyen - added menu option to fix Outlook issues on new computer caused by the registry restore

# Define a function for each menu option

#done add desktop/P drive backup tool added 4/9
#to do add regedit onenote transfer
#to do, add a euvlation code so you dont have to run the regedit as admin.
#to do convert the OneNote shortcuts into ONENOTE:URL shortucts, so you dont have to open edge. added 4/9
#To do fix the auto cleanup of "C:\Temp\LaptopTransferBackups"
#To do fix the random crashing
#to do fix the crash when onedrive is not correctly loaded.
#auto open outlook, Auto open teams, Auto install teams to temp as needed
# users with .499 still not working ex"C:\Users\dhatcher.ALDI-499\OneDrive - ALDI-HOFER" DOne 9/4/24
# to do add possible sync quick pin checker to check health of quick pins
# copy downloads files DONE
# get email signature copyies done
# get computer info in a txt done
#restore e signatures done
#DATA offline mode IE. 9 to run from the temp folder done


####################################################################
# This section is used to grab onenote notebooks paths and names
# Set up some required variables

$onApp = $null                          # Variable to store the OneNote application object
$strPages = ''                          # Variable to store the XML representation of OneNote pages
$xmlPageDoc = New-Object System.Xml.XmlDocument    # XML document object to load and parse OneNote pages
$schema = $null                          # Variable to store the schema based on the OneNote version
$workingFolderPath= "C:\Temp\LaptopTransferBackups"

# example of a bad onedrive item "C:\Users\lindeman.ALDI-499\OneDrive - ALDI-HOFER"

function Get-OneDriveLocation {
    [CmdletBinding()]
    param()

    # Define the potential OneDrive paths
    $oneDrivePath1 = "C:\Users\$($env:USERNAME)\OneDrive - ALDI-HOFER"
    $oneDrivePath2 = "C:\Users\$($env:USERNAME).ALDI-499\OneDrive - ALDI-HOFER"
    $oneDrivePath3 = "C:\Users\$($env:USERNAME).ALDI-499\OneDrive - ALDI DX"
    
    # Check which path exists and return it
    if (Test-Path $oneDrivePath3) {
        return $oneDrivePath3
    }
    if (Test-Path $oneDrivePath2) {
        return $oneDrivePath2
    } elseif (Test-Path $oneDrivePath1) {
        return $oneDrivePath1
    } else {
        Write-Error "OneDrive location not found."
        return $null
    }
}

function Get-FDriveLocation {

    $FDriveLocation = "F:\usrnew\For IT Support\Laptop Swap Script\Backups"

    if(Test-Path $FDriveLocation){
        return $FDriveLocation
    }
    else {
        return $null
    }
}

# Call the function and capture the result in a variable
$OneDriveLocation = Get-OneDriveLocation
$FDriveLocation = Get-FDriveLocation
Function Start-ONApp {
    Param()
    #Write-Host "getting the OneNote application object"
    if (-not $script:onApp) {             # Check if the OneNote application object is not already created
        try {
            #Write-Host "onApp not found"

            # Retrieve the Microsoft.Office.Interop.OneNote assembly from the GAC
            $interOp = Get-Item $env:WinDir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15*\*
            
            #Write-Host "Interop Assembly found at: $($interOp.FullName)"

            # Load the assembly and create the OneNote application object
            Add-Type -LiteralPath $interOp.FullName
            $script:onApp = New-Object Microsoft.Office.Interop.OneNote.ApplicationClass
        } catch [System.Runtime.InteropServices.COMException] {
            Write-Error "Unable to create COM Object - is OneNote installed?"
            Break
        }

        $script:xmlNs = New-Object System.Xml.XmlNamespaceManager($xmlPageDoc.NameTable)
        $onProcess = Get-Process onenote        # Get the OneNote process information
        $onVersion = $onProcess.ProductVersion.Split(".")   # Extract the OneNote version from the process information
        Write-Host "OneNote version $($onVersion[0]) detected"

        # Set the appropriate schema based on the OneNote version
        $script:schema = "http://schemas.microsoft.com/office/onenote/2013/onenote"
        $xmlNs.AddNamespace("one", $schema)      # Add the schema namespace to the XML namespace manager
    } else {
        Write-Host "onApp found"
        #$message = $onApp.GetType()
        #Write-Host $message
    }
}

Function Get-ONHierarchy {
    $onApp.getHierarchy($null, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, [ref] $strPages)
    $xmlPageDoc.LoadXML($strPages)  # Load the XML representation of OneNote pages into the XML document object
}

Function Get-ONNoteBooks {
    $xmlNoteBooks = $xmlPageDoc.SelectNodes("//one:Notebook", $xmlNs)   # Select all OneNote notebooks from the XML document
    
    # Iterate through each notebook node and extract name and path
    $notebooks = foreach ($notebook in $xmlNoteBooks) {
        $name = $notebook.GetAttribute("name")
        $path = $notebook.GetAttribute("path")
        
        # Create a custom object with notebook information
        [PSCustomObject]@{
            Name = $name
            Path = $path
        }
    }
    $notebooks
}

# Runs the above functions in one call
function Get-OneNoteNotebooks {
    param(
        [string]$OutputFilePath = "C:\Temp\LaptopTransferBackups\OneNoteBooks.json"
    )

    Start-ONApp                                     # Ensure that the OneNote application object is created
    Get-ONHierarchy                                 # Get the OneNote hierarchy (pages)
    $notebooks = Get-ONNoteBooks                    # Get the OneNote notebooks
    
    <# # Output the notebook information line by line, just diplay stuff no real work done here
    foreach ($notebook in $notebooks) {
        Write-Host "Notebook Name: $($notebook.Name)"
        Write-Host "Notebook Path: $($notebook.Path)"
    } #>
    
    # Write the $notebooks xml to the $OutputFilePath output file
    if (!(Test-Path (Split-Path -Path $OutputFilePath))) {
        New-Item -ItemType Directory -Path (Split-Path -Path $OutputFilePath) | Out-Null
    }
    Write-Host ""
    $notebooks | Format-Table
    Write-Host ""
    Write-Host "saving notebooks to $OutputFilePath"

    foreach ($notebook in $Notebooks) {
        $name = $notebook.Name
        Write-Host $name
        Start-Sleep -Milliseconds 300
    }

    $notebooks | ConvertTo-Json | Out-File -FilePath $OutputFilePath

    # Output the count
    Write-Host "The file '$OutputFilePath' has recorded" $notebooks.Count "Notebooks."
}

function Get-OneNoteReg {
    # Define the registry path you want to export
    $registryPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\OneNote\OpenNotebooks"  # Change as needed
    
    # Define the output registry file path
    $outputRegFilePath = "C:\Temp\LaptopTransferBackups\OneNoteReg\OneNoteNotebooks.reg"  # Change as needed

    # Define the output file path 
    $outputFilePath = "C:\Temp\LaptopTransferBackups\OneNoteReg\"

    # Check if the source file exists
    Write-Output "The source file $outputFilePath was not found."
   
    # Create the shortcuts folder if it doesn't exist
    if (-not (Test-Path $outputFilePath)) {
       New-Item -ItemType Directory -Path $outputFilePath | Out-Null
    }

    # Export the registry key
    reg export $registryPath $outputRegFilePath /y

    # Confirm completion
    if (Test-Path $outputRegFilePath) {
        Write-Host "Registry key exported successfully to $outputRegFilePath"
    } else {
        Write-Host "Failed to export the registry key."
    }

}

function Get-OutlookReg {
    param(
        [string]$regFileName 
    )
    
    # Define the registry path you want to export
    $registryPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles"  # Change as needed
    
    # Define the output registry file path
    $outputRegFilePath = "C:\Temp\LaptopTransferBackups\OutlookReg\$regFileName.reg"  # Change as needed

    # Define the output file path 
    $outputFilePath = "C:\Temp\LaptopTransferBackups\OutlookReg\"

    
   
    # Create the shortcuts folder if it doesn't exist
    if (-not (Test-Path $outputFilePath)) {
        Write-Output "The source file $outputFilePath was not found."
        New-Item -ItemType Directory -Path $outputFilePath | Out-Null
    }

    # Export the registry key
    reg export $registryPath $outputRegFilePath /y

    # Confirm completion
    if (Test-Path $outputRegFilePath) {
        Write-Host "Registry key exported successfully to $outputRegFilePath"
    } else {
        Write-Host "Failed to export the registry key."
    }

}

function Get-ComputerInfo {
    # Create an ordered dictionary to store the information
    $info = [ordered]@{}

    # Device Name
    $info['Device Name'] = $env:COMPUTERNAME

    # Operating System
    $info['Operating System'] = (Get-CimInstance Win32_OperatingSystem).Caption

    # RAM Amount (in GB)
    $info['RAM (GB)'] = "{0:N2}" -f ((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB)

    # Storage Size (Total size and free space of drives)
    $info['Storage'] = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object { 
        "$($_.DeviceID): $([math]::Round($_.FreeSpace / 1GB, 2))GB free of $([math]::Round($_.Size / 1GB, 2))GB" 
    }

    # Model of the Device
    $info['Model'] = (Get-CimInstance Win32_ComputerSystem).Model

    # BIOS Version
    $info['BIOS Version'] = (Get-CimInstance Win32_BIOS).SMBIOSBIOSVersion

    # Dell Service Tag (ST Number)
    $info['Dell Service Tag'] = (Get-CimInstance Win32_BIOS).SerialNumber

    # Output to console (optional)
    $info.GetEnumerator() | ForEach-Object {
        Write-Output "$($_.Key): $($_.Value)"
    }

    # JSON output file path
    $OutputFilePath = "C:\Temp\LaptopTransferBackups\ComputerInfo.json"  # Specify your desired JSON output path

    # Ensure the directory exists
    if (-not (Test-Path -Path (Split-Path $OutputFilePath -Parent))) {
        New-Item -Path (Split-Path $OutputFilePath -Parent) -ItemType Directory | Out-Null
    }

    # Convert the info hashtable to JSON format
    $json = $info | ConvertTo-Json -Depth 3

    # Export to the JSON file
    Set-Content -Path $OutputFilePath -Value $json

    Write-Output "Computer information exported to: $OutputFilePath"
}
function Open-PathsFromFile {
    # Set the paths
    $sourcePath = "C:\Temp\LaptopTransferBackups\OneNoteBooks.json"
    $shortcutsPath = "C:\Temp\LaptopTransferBackups\OneNoteBooks_shortcuts"

    # Check if the source file exists
    if (-not (Test-Path $sourcePath)) {
        Write-Error "The source file $sourcePath was not found."
        Exit 1
    }

    # Create the shortcuts folder if it doesn't exist
    if (-not (Test-Path $shortcutsPath)) {
        New-Item -ItemType Directory -Path $shortcutsPath | Out-Null
    }

    # Read the JSON content from the file
    $jsonContent = Get-Content $sourcePath -Raw | ConvertFrom-Json

    # Loop through each object and create shortcuts
    $shortcutCount = 0
    foreach ($item in $jsonContent) {
        $path = $item.path

        # Check if the path is a website
        if ($path.StartsWith("http") -or $path.StartsWith("https")) {
            # Add "onenote:" prefix to the URL
            $path = "onenote:" + $path
            
            # Create a shortcut to the website
            Write-Host "Creating Website shortcut for $path"
            $shortcutPath = "$shortcutsPath\$($item.name).url"
            $wshShell = New-Object -ComObject WScript.Shell
            $shortcut = $wshShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $path
            $shortcut.Save()
            $shortcutCount++
        } else {
            # Check if the path is a file path
            if (Test-Path $path) {
                # Find the file named "Open Notebook.onetoc2" and create a shortcut to it
                $notebookPath = Get-ChildItem $path -Recurse -Filter "Open Notebook.onetoc2" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
                if ($notebookPath) {
                    Write-Host "Creating File shortcut for $notebookPath"
                    # Get the folder name and create the shortcut name based on that
                    $folderName = Split-Path -Parent $notebookPath | Split-Path -Leaf
                    $shortcutPath = "$shortcutsPath\$folderName - Open Notebook.lnk"
                    $wshShell = New-Object -ComObject WScript.Shell
                    $shortcut = $wshShell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $notebookPath
                    $shortcut.Save()
                    $shortcutCount++
                } else {
                    Write-Warning "Failed to find 'Open Notebook.onetoc2' file in path: $path"
                }
            }
        }
    }
    # Open the shortcuts folder in File Explorer

    $jsonObjectCount = $jsonContent.Count
    Write-Host "Number of objects in the JSON file: $jsonObjectCount"
    Write-Host "Number of shortcuts created:        $shortcutCount"
    Invoke-Item $shortcutsPath
}
####################################################################
function Compare_Notebooks {

    #Starting OneNote for the first time
    Write-Host "Starting OneNote in 5 -Seconds"
    Start-Sleep -Seconds 5
    Start-Process "onenote.exe"
    Write-Host "Waiting 10 -Seconds for Onenote to open"
    Start-Sleep -Seconds 10

    # will re-read and let you know which note books are open. THen compare to the Old list to see whats missing
    Get-OneNoteNotebooks -OutputFilePath "C:\Temp\LaptopTransferBackups\OneNoteBooksCompare.json"

        $path1 = "C:\Temp\LaptopTransferBackups\OneNoteBooks.json"
        $path2 = "C:\Temp\LaptopTransferBackups\OneNoteBooksCompare.json"
        
            # Read the contents of the JSON files
            $json1 = Get-Content -Path $Path1 -Raw | ConvertFrom-Json
            $json2 = Get-Content -Path $Path2 -Raw | ConvertFrom-Json
        
            # Get the names of the objects in each JSON file
            $names1 = $json1 | ForEach-Object { $_.name }
            $names2 = $json2 | ForEach-Object { $_.name }
        
            # Compare the objects in the two JSON files
            $missingObjects = Compare-Object -ReferenceObject $names1 -DifferenceObject $names2 |
                Where-Object { $_.SideIndicator -eq '<=' } |
                ForEach-Object { $_.InputObject }
        
            # Notify the user if any objects are missing
            if ($missingObjects) {
                Write-Host "The following objects are missing from 'OneNoteBooksCompare.json':"
                $missingObjects | ForEach-Object {
                    $name = $_
                    $missingObject = $json1 | Where-Object { $_.name -eq $name }
                    Write-Host "- Name: $($missingObject.name), Path: $($missingObject.path)"
                }
            } else {
                Write-Host "All objects are present in 'OneNoteBooksCompare.json'."
            }
    }

function Test-OneDriveConnection {

    $OnedriveInstallStaus = $false # by defult this should be false, and only true if it is setup and signed in.
    if (Test-Path $OneDriveLocation) {
        Write-Host "OneDrive folder Found. located at $OneDriveLocation"
        $OnedriveInstallStaus = $true
    }
    else {
        Write-Warning "OneDrive Is not signed into. Due to Windows 11 changes, please have them Sign in to transfer. type Y to check this operation agian." -WarningAction Inquire
        open "C:\Users\$env:USERNAME\AppData\Local\Microsoft\OneDrive\OneDrive.exe"
        Check-OneDrive
    }

}

function Test-FDriveConnection {
    
    $FDrivePresent = $false #default is false to check if computer is a divisional laptop or not
    
    if (Test-Path $FDriveLocation){
        Write-Host "F Drive Folder Found. Located at $FDriveLocation"
        $FDrivePresent = $true
    }
    else{
        Write-Host "No F Drive back up location. Proceeding without using F Drive backup."
    }
}

function Backup-DownloadsFiles {
    $DestinationPath = "C:\Temp\LaptopTransferBackups\DownloadFiles"

    # Get the current user's Downloads path
    $DownloadPath = "$env:USERPROFILE\Downloads"

    # Calculate the total size of files in the Downloads folder
    $totalSizeBytes = (Get-ChildItem -Path $DownloadPath -Recurse | Measure-Object -Property Length -Sum).Sum
    $totalSizeGB = [math]::Round($totalSizeBytes / 1GB, 2)

    if ($totalSizeGB -gt 1)
    {
        # Prompt the user to confirm if they want to proceed, including the expected size
        $confirmation = $null
        $promptMessage = " ${totalSizeGB}GB - Do you want to backup Downloads files? (y/n)"
        $timeout = 999  # Timeout in seconds

        # Start a timer to default to "Yes" after timeout
        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        while (-not $confirmation) {
            if ($timer.Elapsed.TotalSeconds -ge $timeout) {
                # Default to "Yes" if timeout reached
                $confirmation = 'y'
                break
            }

            $userInput = Read-Host -Prompt $promptMessage

            if ($userInput -eq 'y' -or $userInput -eq 'n') {
                $confirmation = $userInput
            }
        }

        if ($confirmation -eq 'n') {
            Write-Host "Backup operation aborted."
            return
        }
    }

    # Check if the directory exists
    if (-not (Test-Path -Path $DestinationPath)) {
        # Create the directory
        New-Item -Path $DestinationPath -ItemType Directory
        Write-Host "Directory created at $DestinationPath"
    } else {
        Write-Host "Directory already exists at $DestinationPath"
    }

    # Construct the destination zip file path
    $zipFilePath = Join-Path -Path $DestinationPath -ChildPath "downloadFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
    Write-Host "Compressing $DownloadPath to $zipFilePath"

    try {
        # Compress the Downloads contents into a zip file
        Compress-Archive -Path "$DownloadPath\*" -DestinationPath $zipFilePath -Force -CompressionLevel "Fastest"

        Write-Output "Download files backed up successfully to: $zipFilePath"
    }
    catch {
        Write-Error "Error compressing Download files: $_"
    }
}

function Find-BackupDirectory {
    $backupPath = "C:\Temp\LaptopTransferBackups"
    Write-Host = "checking status of folder" $backupPath

    if (Test-Path -Path $backupPath -PathType Container) {
        Write-Host "The backup directory $backupPath exists."
        Write-Host "It is recommended to remove this directory before proceeding."

        do {
            $choice = Read-Host -Prompt "Do you want to (D)elete the directory, (S)kip, or (Q)uit? [D/S/Q]"
            $choice = $choice.Trim().ToUpper()

            switch ($choice) {
                'D' {
                    try {
                        # Delete the backup directory
                        Remove-Item -Path $backupPath -Recurse -Force
                        Write-Host "Directory $backupPath has been successfully deleted."
                    }
                    catch {
                        Write-Host "Error deleting directory: $_"
                    }
                    break
                }
                'S' {
                    Write-Host "Skipping removal of the directory."
                    break
                }
                'Q' {
                    Write-Host "Exiting..."
                    exit
                }
                default {
                    Write-Host "Invalid choice. Please choose 'D' to delete, 'S' to skip, or 'Q' to quit."
                }
            }
        } while ($choice -ne 'D' -and $choice -ne 'S' -and $choice -ne 'Q')
    }
    else {
        Write-Host "The backup directory $backupPath does not exist."
    }
}

function Backup-ToOneDrive {

    $sourceFolderPath = "C:\Temp\LaptopTransferBackups"
    $destinationFolderPath = "$OneDriveLocation"
    
    # Check if OneDrive folder exists
    if (-not (Test-Path $destinationFolderPath)) {
        Write-Host "OneDrive folder not found! This should not happen, Exiting"
        # Add any additional actions or error handling here
        Pause
        exit
    }
    
    # Generate a unique timestamped folder name for the backup
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFolderName = "Backup_$timestamp"
    
    # Create the backup folder in OneDrive
    $backupFolderPath = Join-Path -Path $destinationFolderPath -ChildPath $backupFolderName
    New-Item -ItemType Directory -Path $backupFolderPath | Out-Null
    
    # Copy files from source folder to the backup folder
    Copy-Item -Path "$sourceFolderPath\*" -Destination $backupFolderPath -Recurse
    
    Write-Host "Backup completed successfully!"
    Write-Host "Files saved to! $destinationFolderPath\$backupFolderName"
    

}

function Backup-ToFDrive {

    $sourceFolderPath = "C:\Temp\LaptopTransferBackups"
    $destinationFolderPath = "$FDriveLocation"
    
    # Check if destination folder exists
    if (-not (Test-Path $destinationFolderPath)) {
        Write-Host "F Drive folder not found! User either a divisional employee or had a transfer/name change."
        Pause
        exit
    }
    
    # Generate a unique timestamped folder name for the backup
    $backupFolderName = "Backup_$env:USERNAME"
    
    # Create the backup folder in the F Drive
    $backupFolderPath = Join-Path -Path $destinationFolderPath -ChildPath $backupFolderName
    New-Item -ItemType Directory -Path $backupFolderPath -Force | Out-Null
    
    # Define the folders to exclude
    $excludedFolders = @("C:\Temp\LaptopTransferBackups\DownloadFiles", "C:\Temp\LaptopTransferBackups\Wallpaper")

   # Get files to copy
    $filesToCopy = Get-ChildItem -Path $sourceFolderPath -Recurse | Where-Object{
        # Check if the item is a file
        if ($_.PSIsContainer -eq $false) {
            # Exclude files within the excluded folders
            $fileFolder = Split-Path -Path $_.FullName -Parent
            $excluded = $excludedFolders | Where-Object { $fileFolder -like "$_" }
            -not $excluded
        }
    } 

    # Check if any files are found to copy
    if ($filesToCopy.Count -eq 0) {
        Write-Host "No files to copy - check your exclusion criteria and source folder."
        return
    }

    foreach ($file in $filesToCopy) {
        # Determine the destination path
        $destinationPath = $file.FullName -replace [regex]::Escape($sourceFolderPath), $backupFolderPath
        
        # Ensure the destination directory exists
        $destinationDir = Split-Path -Path $destinationPath -Parent
        if (-not (Test-Path -Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
        }
        
        # Copy the file to the destination
        Copy-Item -Path $file.FullName -Destination $destinationPath -Force
    }
    
    Write-Host "Backup completed successfully!"
    Write-Host "Files saved to: $backupFolderPath"
}


function Select-OneDriveBackup {
    $destinationFolderPath = "$OneDriveLocation"
    
    # Check if OneDrive folder exists
    if (-not (Test-Path $destinationFolderPath)) {
        Write-Host "OneDrive folder not found! please make sure the onedrive files are logged in - script will exit"
        # Add any additional actions or error handling here
        pause
        exit
    }

    # Get a list of all backup folders in OneDrive
    $backupFolders = Get-ChildItem -Path $destinationFolderPath | Where-Object { $_.PSIsContainer -and $_.Name -match '^Backup_\d{8}_\d{6}$' } | Sort-Object LastWriteTime -Descending
    
    # Check if there are any backups
    if ($backupFolders.Count -eq 0) {
        Write-Host "No backups found in $destinationFolderPath."
        Write-Host "OneDrive backup not found! please make sure the onedrive files are syncing- script will exit"
        # Add any additional actions or error handling here
        Pause
        exit
    }

    # Display a list of backup folders with their timestamps in a readable format
    Write-Host "Available Backups:"
    for ($i = 0; $i -lt $backupFolders.Count; $i++) {
        $timestamp = $backupFolders[$i].Name -replace 'Backup_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})', '$4:$5:$6 $1-$2-$3'
        $dateTime = Get-Date $timestamp -Format "HH:mm:ss yyyy-MM-dd"
        Write-Host "$i. $dateTime"
    }

    # Prompt the user to select a backup (default to the latest backup)
    $selectedIndex = Read-Host "Select a backup (default is 0 for the latest)" -Default 0

    # Get the selected backup folder path
    $selectedBackup = $backupFolders[$selectedIndex]
    $selectedBackupPath = Join-Path -Path $destinationFolderPath -ChildPath $selectedBackup.Name

    # Specify the destination folder
    $destinationFolder = "C:\Temp\LaptopTransferBackups"

    if (-not (Test-Path(Join-Path -Path $selectedBackupPath -ChildPath "LaptopTransferBackups"))){
        New-Item -ItemType Directory -Path "C:\Temp\LaptopTransferBackups" | Out-Null
         # Copy only the contents of the selected backup folder to the destination folder
    }
    else {
        $selectedBackupPath = Join-Path -Path $selectedBackupPath -ChildPath "LaptopTransferBackups"

    }
    # Get all items in the selected backup folder, excluding Desktop and Downloads
    $itemsToCopy = Get-ChildItem -Path $selectedBackupPath -Recurse | Where-Object {
        # Exclude if the top-level folder name is Desktop or Downloads
        ($_.PSIsContainer -and ($_.FullName -replace [regex]::Escape($selectedBackupPath), '') -split '[\\/]')[1] -notin 'Desktop','Downloads' -or -not $_.PSIsContainer
    }

    foreach ($item in $itemsToCopy) {
        $relativePath = $item.FullName.Substring($selectedBackupPath.Length)
        $destPath = Join-Path -Path $destinationFolder -ChildPath $relativePath
        if ($item.PSIsContainer) {
            if (-not (Test-Path $destPath)) {
                New-Item -ItemType Directory -Path $destPath | Out-Null
            }
        } else {
            Copy-Item -Path $item.FullName -Destination $destPath
        }
    }
    if (Test-Path "$FDriveLocation\Backup_$env:USERNAME"){
        Remove-Item -Path "$FDriveLocation\Backup_$env:USERNAME" -Recurse -Force
    }

    # Return the selected backup path
    return $selectedBackupPath
}

function Select-FDriveBackup{

    $destinationFolderPath = "$FDriveLocation\Backup_$env:USERNAME"
    
    # Check if F Drive folder exists
    if (-not (Test-Path $destinationFolderPath)) {
        Write-Host "F Drive folder not found! Please make sure the  F Drive files are present and match the current username - script will exit"
        # Add any additional actions or error handling here
        pause
        exit
    }


    # Specify the destination folder
    $destinationFolder = "C:\Temp\LaptopTransferBackups"

    if (-not (Test-Path(Join-Path -Path $selectedBackupPath -ChildPath "LaptopTransferBackups"))){
        New-Item -ItemType Directory -Path "C:\Temp\LaptopTransferBackups" | Out-Null
    }

    # Copy only the contents of the selected backup folder to the destination folder
    Copy-Item -Path "$destinationFolderPath\*" -Destination $destinationFolder -Recurse

    Remove-Item -Path "$destinationFolderPath" -Recurse -Force


}


function Backup-QuickAccess {
    # Set the source and destination paths
    $sourcePath = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"
    $destinationPath = "C:\Temp\LaptopTransferBackups\QuickAcessBK"

    # Check if the source directory exists
    if (Test-Path $sourcePath) {
        # Create the destination directory if it doesn't exist
        if (!(Test-Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath | Out-Null
        }

        # Get all the files in the source directory
        $files = Get-ChildItem $sourcePath

        # Initialize progress bar
        $progressParams = @{
            Activity = "Backing up Quick Access Files"
            Status = "Initializing..."
            PercentComplete = 0
        }
        $progress = $null

        # Copy each file to the destination directory
        foreach ($file in $files) {
            $progressParams.PercentComplete = (($files.IndexOf($file) + 1) / $files.Count) * 100
            $progressParams.Status = "$($file.Name) is backing up"
            $progress = Write-Progress @progressParams -Id 1 -ParentId 0

            $destinationFile = Join-Path $destinationPath $file.Name
            Copy-Item $file.FullName $destinationFile -Force
        }

        # Complete the progress bar
        $progressParams.PercentComplete = 100
        $progressParams.Status = "Quick Link Files backed up to $destinationPath."
        Write-Progress @progressParams -Id 1 -ParentId 0 -Completed

        Write-Host "Quick Link Files backed up to $destinationPath."
    }
    else {
        Write-Host "Source directory not found."
    }
}

function Restore-QuickAccess {
    # Set the source and destination paths
    $sourcePath = "C:\Temp\LaptopTransferBackups\QuickAcessBK"
    $destinationPath = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations"

    # Check if the source directory exists
    if (Test-Path $sourcePath) {
            robocopy /is $sourcePath $destinationPath /MT
            Write-Host "Files restored to $destinationPath."
            # Restart File Explorer to refresh its view
            Stop-Process -Name explorer -Force
        }

    else {
        Write-Warning "Source directory not found.! Please Run Quick access Backup tool and Sync"  -WarningAction Inquire

    }
}

function Copy-EmailSignatures {
    # Define source and destination paths
    $sourcePath = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Signatures"
    $destinationPath = "C:\Temp\LaptopTransferBackups\EmailSignatures"

    Write-Host "Starting the copy process for email signatures..."
    Write-Host "Source path: $sourcePath"
    Write-Host "Destination path: $destinationPath"

    # Check if the source directory exists
    if (-not (Test-Path $sourcePath)) {
        Write-Host "Source path does not exist: $sourcePath" -ForegroundColor Red
        return
    }

    # Create the destination folder if it doesn't exist
    if (-not (Test-Path $destinationPath)) {
        Write-Host "Destination path does not exist. Creating directory: $destinationPath"
        New-Item -Path $destinationPath -ItemType Directory
    } else {
        Write-Host "Destination path already exists: $destinationPath"
    }

    # Use RoboCopy to copy the files
    robocopy $sourcePath $destinationPath /E /COPY:DAT

}

function restore-EmailSignatures {
    # Define source and destination paths
    $destinationPath = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Signatures"
    $sourcePath = "C:\Temp\LaptopTransferBackups\EmailSignatures"

    Write-Host "Starting the copy process for email signatures..."
    Write-Host "Source path: $sourcePath"
    Write-Host "Destination path: $destinationPath"

    # Check if the source directory exists
    if (-not (Test-Path $sourcePath)) {
        Write-Host "Source path does not exist: $sourcePath" -ForegroundColor Red
        return
    }

    # Create the destination folder if it doesn't exist
    if (-not (Test-Path $destinationPath)) {
        Write-Host "Destination path does not exist. Creating directory: $destinationPath"
        New-Item -Path $destinationPath -ItemType Directory
    } else {
        Write-Host "Destination path already exists: $destinationPath"
    }

    # Use RoboCopy to copy the files
    robocopy $sourcePath $destinationPath /E /COPY:DAT

}

function Restore-OneNoteNotebooks {

    # Define the path to the .reg file stored in OneDrive
    $regFilePath = "C:\Temp\LaptopTransferBackups\OneNoteReg\OneNoteNotebooks.reg"  # Change as needed

    #Kill OneNote and Send to OneNote Tool processes
    # Get the OneNote process or Send to OneNote tool process and stop them
    Get-Process -Name "OneNote" -ErrorAction SilentlyContinue | Stop-Process -Force
    Get-Process -Name "ONENOTEM" -ErrorAction SilentlyContinue | Stop-Process -Force

    Write-Host "Waiting 10 -Seconds for Onenote to close"
    Start-Sleep -Seconds 10

    # Check if the .reg file exists
    if (Test-Path $regFilePath) {
        # Import the registry file
        reg import $regFilePath

        # Confirm completion
        Write-Host "Registry file imported successfully from $regFilePath"
    } else {
        Write-Host "The specified registry file does not exist: $regFilePath"
        
    }
}

# Function to check for the registry key
function Wait-ForRegistryKey {
    param (
        [string]$path,
        [int]$timeoutInSeconds
    )

    $startTime = Get-Date
    $keyFound = $false

    while (-not $keyFound) {
        # Check if the registry key exists
        if (Test-Path $path) {
            Write-Output "Registry key found: $path"
            $keyFound = $true
        } else {
            # Wait for 1 second before checking again
            Start-Sleep -Seconds 1

            # Check if the timeout has been reached
            $elapsedTime = (Get-Date) - $startTime
            if ($elapsedTime.TotalSeconds -ge $timeoutInSeconds) {
                Write-Output "Timeout reached. The registry key was not found within $timeoutInSeconds seconds."
                break
            }
        }
    }
}

function Restore-OutlookMailboxes {

    # Define the path to the .reg file stored in Temp folder
    $regFilePath = "C:\Temp\LaptopTransferBackups\OutlookReg\OldPcOutlook.reg"
    # Define the path to check for Outlook registry entry to verify Outlook has been opened
    $registryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

    Wait-ForRegistryKey -path $registryPath -timeoutInSeconds 999

    Get-OutlookReg -regFileName "NewPcBackup"

    try {
        # Get the Outlook process
        $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue

        # If Outlook is running, stop it
        if ($outlookProcess) {
            Write-Output "Stopping Outlook..."
            Stop-Process -Name "OUTLOOK" -Force
            Write-Output "Outlook has been stopped."
        } else {
            Write-Output "Outlook is not running."
        }
    } catch {
        Write-Output "An error occurred: $_"
    }

    Write-Host "Waiting 10 -Seconds for Outlook to close"
    Start-Sleep -Seconds 10

    # Check if the .reg file exists
    if (Test-Path $regFilePath) {
        # Import the registry file
        reg import $regFilePath

        # Confirm completion
        Write-Host "Registry file imported successfully from $regFilePath"
    } else {
        Write-Host "The specified registry file does not exist: $regFilePath"
        
    }
}

function Fix-OutlookRegistry {

    # Define the path to the .reg file stored in Temp folder
    $regFilePath = "C:\Temp\LaptopTransferBackups\OutlookReg\NewPcBackup.reg"

    try {
        # Get the Outlook process
        $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue

        # If Outlook is running, stop it
        if ($outlookProcess) {
            Write-Output "Stopping Outlook..."
            Stop-Process -Name "OUTLOOK" -Force
            Write-Output "Outlook has been stopped."
        } else {
            Write-Output "Outlook is not running."
        }
    } catch {
        Write-Output "An error occurred: $_"
    }

    Write-Host "Waiting 10 -Seconds for Outlook to close"
    Start-Sleep -Seconds 10

    # Check if the .reg file exists
    if (Test-Path $regFilePath) {
        # Import the registry file
        reg import $regFilePath

        # Confirm completion
        Write-Host "Registry file imported successfully from $regFilePath"
    } else {
        Write-Host "The specified registry file does not exist: $regFilePath"
        
    }
}

#Start and stop log capture
function Write-OutputToLog {
    param (
        [bool]$status
    )

    # Get current date and time to use in the filename
    $currentDate = Get-Date -Format "yyyy-MM-dd"
    $filePath = "C:\Temp\LaptopTransferBackups\Logs\Log_$currentDate.txt"
    $outputFilePath = "C:\Temp\LaptopTransferBackups\Logs"
    
    

    # Check if the file/directory exists, create it if it doesn't
    if (-Not (Test-Path -Path $filePath)) {
        if (-Not (Test-Path -Path $outputFilePath)) {
            New-Item -ItemType Directory -Path $outputFilePath | Out-Null
        }
        New-Item -Path $filePath -ItemType File | Out-Null
    }
    #starts log capture if status is true
    if ($status){
        Start-Transcript -Path $filePath
    }
    #ends log capture if status is false
    else {
        Stop-Transcript
    }
}

#Function to get the file path tree of a folder and send it to a .txt file
function Get-Tree {
    param (
        # Define the directory you want to save the path tree for
        [string]$directory,
        [string]$treeTarget
    )

    # Define the output file
    $filePath = "C:\Temp\LaptopTransferBackups\Trees\$treeTarget.txt" # Change this to your desired output file path
    $outputFile = "C:\Temp\LaptopTransferBackups\Trees"

    #Checks if the file path to save the tree to exists or not and creates it if it doesn't
    if (-Not (Test-Path -Path $filePath)) {
        if (-Not (Test-Path -Path $outputFile)) {
            New-Item -ItemType Directory -Path $outputFile | Out-Null
        }
        New-Item -Path $filePath -ItemType File | Out-Null
    }
    
    #Gets file path tree and saves it to a .txt file
    tree $directory /f > $filePath
}

# Function to get the current wallpaper
function Get-CurrentWallpaper {
    # Query the current wallpaper path from the registry
    $registryPath = "HKCU:\Control Panel\Desktop"
    $wallpaperPath = (Get-ItemProperty -Path $registryPath -Name Wallpaper).Wallpaper

    # Return the wallpaper path
    return $wallpaperPath
}

# Function to backup the wallpaper
function Backup-Wallpaper {
    # Retrieve the current wallpaper path
    $currentWallpaper = Get-CurrentWallpaper
    $backupPath = "C:\Temp\LaptopTransferBackups\Wallpaper"

    # Create the backup folder if it doesn't exist
    if (-not (Test-Path -Path $backupPath)) {
        Write-Host "The backup folder $backupPath was not found. Creating file path."
        New-Item -ItemType Directory -Path $backupPath
    }

    # Check if the current wallpaper exists and copy it to the backup location
    if ($currentWallpaper -and (Test-Path -Path $currentWallpaper)) {        
        try {
            Copy-Item -Path $currentWallpaper -Destination $backupPath
            Write-Host "Wallpaper backed up successfully to $backupPath"
        } catch {
            Write-Host "Error while copying the wallpaper: $_"
        }
    } else {
        Write-Host "Current wallpaper not found or path is incorrect."
    }
}


function Write-Tree{
    $downloadsLocation = "$env:USERPROFILE\Downloads"

    Get-Tree -directory $downloadsLocation -treeTarget "Downloads"
}
function Get-AppList {
    <# To Add Applicaiton
    1. write the name of the app in $applications list
    2. Add the same name to the switch ($app), and add the path into {$path = ""}
    #>
    # to do add a way to Json file store the apps and have them Auto open in Five9s
    # List of applications to check
    $applications = "RedPrairie (MCH)", "Tableau Prep (Server)", "Tableau Desktop", "Spaceman", "Kofax", "Alteryx", "Git", "SSMS (SQL Server Management Studio)", "Anaconda", "Python", "7-Zip", "Notepad++", "Visio", "Adobe Creative Cloud", "Think-Cell", "Visual Studio Code", "keepass", "kerberrose", "JDA Enterprise Client", "Visual Studio"

    # Array to store information about installed applications
    $installedApps = @()

    foreach ($app in $applications) {
        # Get the default installation path for the application
        switch ($app) {
            "RedPrairie (MCH)"          { $path = "C:\Program Files (x86)\RedPrairie\MOCA\client" }
            "Tableau Prep (Server)"     { $path = "C:\Program Files\Tableau Prep Builder" }
            "Tableau Desktop"           { $path = "C:\Program Files\Tableau\Tableau [0-9].[0-9]" }
            "Spaceman"                  { $path = "C:\Program Files\Spaceman" }
            "Kofax"                     { $path = "C:\Program Files (x86)\Kofax\AcrobatConnector" }
            "Alteryx"                   { $path = "C:\Program Files\Alteryx" }
            "Git"                       { $path = "C:\Program Files\Git" }
            "SSMS (SQL Server Management Studio)"      { $path = "C:\Program Files (x86)\Microsoft SQL Server\140\Tools\Binn\ManagementStudio\SSMS.exe" }
            "Anaconda"                  { $path = "C:\Program Files\Anaconda3\python.exe" }
            "Python"                    { $path = "C:\Program Files\Python[0-9][0-9]\python.exe" }
            "7-Zip"                     { $path = "C:\Program Files\7-Zip\7z.exe" }
            "Notepad++"                 { $path = "C:\Program Files\Notepad++\Notepad++.exe" }
            "Visio"                     { $path = "C:\Program Files\Microsoft Office\root\Office[0-9][0-9]\VISIO.EXE" }
            "Adobe Creative Cloud"      { $path = "C:\Program Files (x86)\Adobe\Adobe Creative Cloud\CoreSync\CoreSync.exe" }
            "Think-Cell"                 { $path = "C:\Program Files\think-cell" }
            "Visual Studio Code"        { $path = "C:\Program Files\Microsoft VS Code\Code.exe" }
            "Visual Studio"             { $path = "C:\Program Files\Microsoft VS Code\" }
            "keepass"                   { $path = "C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe" }
            "kerberrose"                { $path = "C:\Program Files (x86)\Kerberos\Kerberos.exe" }
            "JDA Enterprise Client"     { $path = "C:\Program Files (x86)\Kerberos\Kerberos.exe" }
            default                     { Write-Host "Installation path for $app not defined" }
        }

        # Check if the application is installed by looking for its executable file in the default path
        if (Test-Path $path) {
            # Add the installed application to the array with information about its installation status and path
            $installedApps += [PSCustomObject] @{
                Application = $app
                Installed = "Yes"
                Path = $path
            }
        } else { 
            # Add the application to the array with information about its installation status and a "Not found" path
            $installedApps += [PSCustomObject] @{
                Application = $app
                Installed = "No"
                Path = "Not found"
            }
        }
    }

    # Filter the installed applications
    $installedAppsOnly = $installedApps | Where-Object { $_.Installed -eq "Yes" }

    # Display the installed applications as a table
    $installedAppsOnly | Format-Table -AutoSize
}

$Spacer = "#######################################" + "`n"
function Start-Backup {

    Write-Host "STARTING FUll AUTO BACKUP..."
    $Spacer

    Write-Host "Checking If temp files are left"
    Find-BackupDirectory
    $Spacer

    Write-Host "Starting Log Capture"
    Write-OutputToLog -status $true
    $Spacer

    Write-Host "Checking If OneDrive is setup"
    Test-OneDriveConnection
    $Spacer

    Write-Host "Checking for F Drive"
    Test-FDriveConnection
    $Spacer

    Write-Host "Pulling Back'd up computer info"
    Get-ComputerInfo
    $Spacer

    Write-Host "Pulling Backup Email signatures"
    Copy-EmailSignatures
    $Spacer

    Write-Host "Getting Onenotes via assmebly "
    Get-OneNoteNotebooks
    $Spacer
    
    Write-Host "Creating OneNote Shortcuts"
    Open-PathsFromFile
    $Spacer

    Write-Host "Getting Outlook Mailboxes"
    Get-OutlookReg -regFileName "OldPcOutlook"
    $Spacer

    Write-Host "Getting OneNote registry files"
    Get-OneNoteReg
    $Spacer

    Write-Host "Backing Up quick access"
    Backup-QuickAccess
    $Spacer

    Write-Host "Backing Up Download files"
    Backup-DownloadsFiles
    $Spacer
    
    Write-Host "Backing up Wallpaper"
    Backup-Wallpaper
    $Spacer

    Write-Host "Getting File Path Trees"
    Write-Tree
    $Spacer

    Write-Host "Creating Backup File to one drive "
    Backup-ToOneDrive
    $Spacer

    Write-Host "Creating Backup Files to F Drive"
    Backup-ToFDrive
    $Spacer

    Write-Host "Table Creation of all common apps"
    Get-AppList
    $Spacer

    Write-Host "Cleaning UP $workingFolderPath"
    #Remove-Item -Path $workingFolderPath -Recurse -Force
    $Spacer

    Write-Host "FUll AUTO BACKUP DONE"
    $Spacer

    Write-Host "Ending log capture"
    Write-OutputToLog -status $false

}

function Start-Restore{
    Write-Host "STARTING FUll AUTO RESTORE..."
    $Spacer
    
    Write-Host "Starting Log Capture"
    Write-OutputToLog -status $true
    $Spacer

    Write-Host "Starting Outlook"
    Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    $Spacer

    Write-Host "Checking If OneDrive is setup"
    Test-OneDriveConnection
    $Spacer

    Write-Host "Scanning For Onedrive Backups"
    $selectedBackupPath = Select-OneDriveBackup
    Write-Host "Backup selected as $selectedBackupPath"
    $Spacer

    Write-Host "Restoring EmailSignatures files"
    restore-EmailSignatures
    $Spacer

    Write-Host "Restoring QuickAccess"
    $Spacer

    Restore-QuickAccess
    Write-Host "Waiting for file reboot 10 seconds"
    Start-Sleep -Seconds 10
    $Spacer
    
    Write-Host "Seeing If any notebooks are missing"
    Compare_Notebooks 
    $Spacer

    Write-Host "Importing OneNote notebooks registry file"
    Restore-OneNoteNotebooks
    $Spacer

    Write-Host "Restoring Outlook Mailboxes"
    Restore-OutlookMailboxes
    $Spacer

    Write-Host "Cleaning UP $workingFolderPath"
    # Remove-Item -Path $workingFolderPath -Recurse -Force
    $Spacer

    Write-Host "FUll AUTO RESTORE DONE."
    $Spacer

    Write-Host "Ending log capture"
    Write-OutputToLog -status $false
}

function Start-FDriveRestore{
    Write-Host "STARTING FUll AUTO RESTORE..."
    $Spacer
    
    Write-Host "Starting Log Capture"
    Write-OutputToLog -status $true
    $Spacer

    Write-Host "Starting Outlook"
    Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    $Spacer

    Write-Host "Checking If F Drive is setup"
    Test-FDriveConnection
    $Spacer

    Write-Host "Scanning For F Drive Backups"
    Select-FDriveBackup
    $Spacer

    Write-Host "Restoring EmailSignatures files"
    restore-EmailSignatures
    $Spacer

    Write-Host "Restoring QuickAccess"
    $Spacer

    Restore-QuickAccess
    Write-Host "Waiting for file reboot 10 seconds"
    Start-Sleep -Seconds 10
    $Spacer

    Write-Host "Creating OneNote Shortcuts"
    Open-PathsFromFile
    $Spacer
    
    Write-Host "Seeing If any notebooks are missing"
    Compare_Notebooks 
    $Spacer

    Write-Host "Importing OneNote notebooks registry file"
    Restore-OneNoteNotebooks
    $Spacer

    Write-Host "Restoring Outlook Mailboxes"
    Restore-OutlookMailboxes
    $Spacer

    Write-Host "Cleaning UP $workingFolderPath"
    # Remove-Item -Path $workingFolderPath -Recurse -Force
    $Spacer

    Write-Host "FUll AUTO RESTORE DONE."
    $Spacer

    Write-Host "Ending log capture"
    Write-OutputToLog -status $false
}

function Start-OutlooklessRestore{
     Write-Host "STARTING FUll AUTO RESTORE..."
    $Spacer
    
    Write-Host "Starting Log Capture"
    Write-OutputToLog -status $true
    $Spacer


    Write-Host "Checking If OneDrive is setup"
    Test-OneDriveConnection
    $Spacer

    Write-Host "Scanning For Onedrive Backups"
    $selectedBackupPath = Select-OneDriveBackup
    Write-Host "Backup selected as $selectedBackupPath"
    $Spacer

    Write-Host "Restoring EmailSignatures files"
    restore-EmailSignatures
    $Spacer

    Write-Host "Restoring QuickAccess"
    $Spacer

    Restore-QuickAccess
    Write-Host "Waiting for file reboot 10 seconds"
    Start-Sleep -Seconds 10
    $Spacer
    
    Write-Host "Seeing If any notebooks are missing"
    Compare_Notebooks 
    $Spacer

    Write-Host "Importing OneNote notebooks registry file"
    Restore-OneNoteNotebooks
    $Spacer

    Write-Host "Cleaning UP $workingFolderPath"
    # Remove-Item -Path $workingFolderPath -Recurse -Force
    $Spacer

    Write-Host "FUll AUTO RESTORE DONE."
    $Spacer

    Write-Host "Ending log capture"
    Write-OutputToLog -status $false
}

function Start-OfflineRestore{
    Write-Host "STARTING offline AUTO RESTORE..."
    $Spacer

    Write-Host "Checking if path exists"
        # Check if the directory exists
    if (!(Test-Path -Path "C:\Temp\LaptopTransferBackups")) {
        New-Item -ItemType Directory -Path "C:\Temp\LaptopTransferBackups" -Force
        Write-Output "Directory created at C:\Temp\LaptopTransferBackups"
        Write-Host "you have these files in there"
        tree "C:\Temp\LaptopTransferBackups" /f
    } else {
        # If it exists, continue
        Write-Output "Directory already exists at C:\Temp\LaptopTransferBackups"
        Write-Host "you have these files in there"
        tree "C:\Temp\LaptopTransferBackups" /f
    }
    $Spacer

    Write-Host "Starting Log Capture"
    Write-OutputToLog -status $true
    $Spacer

    Write-Warning "You Have to Download Backup_yyyymmdd_tttttt Files from https://asgportal-my.sharepoint.com/my and place in C:\Temp\LaptopTransferBackups. File structure needs to be C:\Temp\LaptopTransferBackups\*FILES*" -WarningAction Inquire
    $Spacer

    Write-Host "Restoring EmailSignatures files"
    restore-EmailSignatures
    $Spacer

    Write-Host "Restoring QuickAccess"
    $Spacer

    Restore-QuickAccess
    Write-Host "Waiting for file reboot 10 seconds"
    Start-Sleep -Seconds 10
    $Spacer

    Write-Host "Creating OneNote Shortcuts"
    Open-PathsFromFile
    $Spacer

    Write-Host "Seeing If any notebooks are missing"
    Compare_Notebooks 
    $Spacer

    Write-Host "Importing OneNote notebooks registry file"
    Restore-OneNoteNotebooks
    $Spacer

    Write-Host "Cleaning UP $workingFolderPath"
    # Remove-Item -Path $workingFolderPath -Recurse -Force
    $Spacer

    Write-Host "FUll AUTO RESTORE DONE."
    $Spacer

    Write-Host "Ending of log capture"
    Write-OutputToLog -status $false
}

# Define the menu options
$MenuOptions = @{
    1 = "Start-Backup"
    2 = "Start-Restore"
    3 = "Start-FDriveRestore"
    4 = "Start-OfflineRestore"
    5 = "Start-OutlooklessRestore"
    6 = "Fix-OutlookRegistry"
    7 = "Get-AppList"
    8 = "Test-OneDriveConnection"
    q = "Quit"
}

# Loop through the menu options until the user quits
while ($true) {
    Write-Host "`n"
    Write-Host "=== MENU ==="
    Write-Host "Please select an option:"
    foreach ($key in $MenuOptions.Keys) {
        Write-Host "$key. $($MenuOptions[$key])"
    }
    $Selection = Read-Host "Enter a number or 'q' to quit"
    if ($Selection -eq "q") {
        break
    }
    $SelectedOption = $MenuOptions[[int]$Selection]
    # Call the function corresponding to the selected option
    & $SelectedOption
}
