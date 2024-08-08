<#
.SYNOPSIS
    Script to flatten a folder and subfolders to a specified depth, and rename files to remove special characters. 
    The root folder is depth 1. Default depth value is 2, where all objects in recursive subfolders are moved to the first root subfolders.

.DESCRIPTION
    This script processes a specified root folder and its subfolders up to a given depth, moving files to a target folder and renaming them to remove special characters. 
    It can also generate a report of expected changes without making any modifications if the WhatIf parameter is set.

.PARAMETER SourceFolder
    Specifies the source folder to target for script processing.

.PARAMETER CopyDestination
    Specifies the destination folder for processed root folder and files.

.PARAMETER ModifyOriginal
    If set, will override CopyDestination and will modifying the original folders and files.

.PARAMETER Depth
    Sets the depth of folder from root. Objects in folders deeper than this level will be moved to the subfolder at this level.

.PARAMETER CreateZip
    After file and folder processing is completed, creates a zip archive of the modified folder in the folder's parent directory.

.PARAMETER CleanupFiles
    Deletes the modified CopyDestination folder or the modified source folder if -ModifyOriginal is enabled.
    Can only be used with the -CreateZip option.

.PARAMETER WhatIf
    If set, will create a report with all expected changes without modifying folders or files.

.PARAMETER NoPrompt
    If set, will run script with command arguments if declared or default values, and process without any interactive prompts.

.PARAMETER Quiet
    If set, only script status and error messages will be printed to the console, not individual rename and move events.

.PARAMETER RegexPattern
    Defines the regex pattern for illegal characters to be removed from file names.

.PARAMETER ConfirmDestructiveProcess
    Required as a safeguard to use -ModifyOriginal

.EXAMPLE
    .\FolderFlattenAndFileRename.ps1

.EXAMPLE
    .\FolderFlattenAndFileRename.ps1 -SourceFolder C:\SourceFolder -CopyDestination C:\DestinationFolder -Depth 3 -CreateZip -CleanupFiles -NoPrompt

.EXAMPLE
    .\FolderFlattenAndFileRename.ps1 -SourceFolder C:\SourceFolder -ModifyOriginal -ConfirmDestructiveProcess -Quiet -CreateZip -NoPrompt

.EXAMPLE
    .\FolderFlattenAndFileRename.ps1 -SourceFolder C:\SourceFolder -CopyDestination C:\DestinationFolder -Depth 1 -WhatIf

.NOTES
    Author: Jeff Cloherty [https://github.com/jeffcloherty/]
    Created: 7/17/2024
    Version: 1.0.1
    Last Updated: 8/8/2024
    Revision History:
    - 1.0.0 Initial version
    - 1.0.1 Added folder handling for SMB network shares.
            Added configuration parameters to log output.
            Updated function names to only use PowerShell approved verbs.
            BUGFIX: Name collision handling when trying to rename or copy a folder.
    - 1.0.2 Added CreateZip option.
            Added CleanupFiles option.
    - 1.0.3 BUGFIX: Accept SMB SourceFolder with '$' character in path. 
#>

### Default Parameters
[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [string]$SourceFolder = $PSScriptRoot,

    [ValidateRange(1, 10)]
    [int]$Depth = 2,

    [switch]$ModifyOriginal,
    [string]$CopyDestination,
    [switch]$CreateZip,
    [switch]$CleanupFiles,

    [switch]$WhatIf,
    [switch]$NoPrompt,
    [switch]$Quiet,
    [switch]$ConfirmDestructiveProcess,

    # Define the regex pattern for illegal characters
    [string]$RegexPattern = '[^a-zA-Z0-9]', # Remove all special characters. Only permit a-z, A-Z and 0-9.
    #[string]$RegexPattern = '[^a-zA-Z0-9.-_]', # Remove special characters except for: - _ .
    [string]$ReplacementCharacter = "_" # Can be blank "" to remove the special characters without replacement.
)

$ParamPropertyLength = ((((Get-Command -Name $MyInvocation.InvocationName).Parameters).Keys | 
    ForEach-Object{$var = Get-Variable -Name $_ -ErrorAction SilentlyContinue; if($var){$var.Name}})| 
    Measure-Object -Maximum -Property Length).Maximum


### Function Definitions

# Function maintains uniform timestamp format for logs and output.
function Get-Timestamp{
    [string](Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

# Function to log actions to a CSV file
function New-LogRecord {
    param (
        [Parameter(Mandatory=$true)]
        [string]$LogType,
        [string]$ActionStatus,
        [string]$TargetObject,
        [string]$Destination,
        [string]$ErrMsg
    )

    # Create a detailed descriptive message
    if ($ErrMsg) {
        $Message = "[$(Get-Timestamp)] [$LogType] action failed on [$TargetObject] with error: $ErrMsg"
    } elseif($LogType -eq "GetParam"){
        $Message = "[$(Get-Timestamp)] [$LogType] $($TargetObject.PadRight($ParamPropertyLength, ' ')) = $Destination"
    }elseif($LogType -eq "SetParam"){
        $Message = "[$(Get-Timestamp)] [$LogType] action completed: $TargetObject = $Destination"
    }elseif($Destination) {
        $Message = "[$(Get-Timestamp)] [$LogType] action completed: [$TargetObject] to [$Destination]"
    } else {
        $Message = "[$(Get-Timestamp)] [$LogType] action completed on [$TargetObject]"
    }

    $logEntry = [PSCustomObject]@{
        TimeStamp = Get-Timestamp
        LogType = $LogType
        ActionStatus = $ActionStatus
        Target = $TargetObject
        Destination = $Destination
        Message = $Message
    }

    $logEntry | Export-Csv -LiteralPath $logFilePath -Append -NoTypeInformation

    # Print to console based on Quiet switch
    if($Quiet){
        if ($LogType -notlike "rename*" -and $LogType -notlike "move*" -and $LogType -notlike "delete*"){
            Write-Output $Message
        }
    }else{
        Write-Output $Message
    }
}

# Function to process the next folder level
function Invoke-ProcessNextFolderLevel {
    param (
        [Parameter(Mandatory=$true)]
        [string]$thisFolderPath,
        [Parameter(Mandatory=$true)]
        [int]$thisFolderDepth
    )
    try {
        if ($thisFolderDepth -lt $Depth) {
            Write-Output "[$(Get-Timestamp)] Folder depth set to $Depth. Skipping level ${thisFolderDepth}: $thisFolderPath"
            Get-ChildItem $thisFolderPath -Directory | 
                ForEach-Object { 
                    Invoke-ProcessNextFolderLevel -thisFolderPath $_.FullName -thisFolderDepth ($thisFolderDepth + 1) 
                }
        } elseif ($thisFolderDepth -eq $Depth) {
            if ((Get-ChildItem $thisFolderPath -Directory).count -eq 0) {
                Write-Output "[$(Get-Timestamp)] No subfolders to process under: $thisFolderPath"
            } else {
                # Move all files in subfolders to this folder
                Write-Output "[$(Get-Timestamp)] Move all files in subfolders to $thisFolderPath"
                Get-ChildItem $thisFolderPath -File -Force -Recurse | 
                    Where-Object { $_.Directory.FullName -ne $thisFolderPath } | 
                    ForEach-Object { Move-File -File $_ -TargetFolder $thisFolderPath }

                # Delete empty subfolders
                do {
                    Write-Output "[$(Get-Timestamp)] Delete all empty subfolders from $thisFolderPath"
                    $deleted = Remove-EmptyFolders -rootPath $thisFolderPath
                } while ($deleted)
            }
        }
    }
    catch {
        New-LogRecord -LogType "ProcessFolder" -ActionStatus "error" -TargetObject $thisFolderPath -ErrMsg $_.Exception.Message
    }
}

# Function to move files when flattening folders
function Move-File {
    param (
        [Parameter(Mandatory=$true)]
        $File,
        [Parameter(Mandatory=$true)]
        [string]$TargetFolder
    )
    try {
        # If file path exists, rename file to be moved.
        if (Test-Path -LiteralPath "$TargetFolder\$($File.Name)") {
            for ($i = 1; $i -le 100; $i++) {
                $fileName = "$($file.BaseName)-$i$($file.Extension)"
                if (-not (Test-Path -LiteralPath "$TargetFolder\$fileName")) {
                    break
                }
            }
        } else {
            $fileName = $File.Name
        }

        # Move file
        if ($WhatIf) {
            Move-Item -LiteralPath $File.FullName -Destination "$TargetFolder\$fileName" -WhatIf
            New-LogRecord -LogType "MoveFile" -ActionStatus "expected" -TargetObject $File.FullName -Destination "$TargetFolder\$fileName"
        } else {
            $result = Move-Item -LiteralPath $File.FullName -Destination "$TargetFolder\$fileName" -PassThru
            New-LogRecord -LogType "MoveFile" -ActionStatus "complete" -TargetObject $File.FullName -Destination $result.FullName
        }
    }
    catch {
        New-LogRecord -LogType "MoveFile" -ActionStatus "error" -TargetObject $File.FullName -Destination "$TargetFolder\$fileName" -ErrMsg $_.Exception.Message
    }
}

# Function to remove special characters from folder names.
function Rename-Folder {
    param (
        [Parameter(Mandatory=$true)]
        $Folder
    )
    try {    
        # Create a new name with illegal characters removed
        $newName = $Folder.Name -replace $RegexPattern, $ReplacementCharacter
        
        # Check if folder with same name already exists
        if (Test-Path -LiteralPath (Join-Path -Path $Folder.Parent -ChildPath $newName)) {
            for ($i = 1; $i -le 100; $i++) {
                $newName = "$newName-$i"
                if (-not (Test-Path -LiteralPath (Join-Path -Path $Folder.Parent -ChildPath $newName))) {
                    break
                }
            }
        }
        
        # Rename the folder if the new name is different
        if ($newName -ne $Folder.Name) {
            if ($WhatIf) {
                Rename-Item -LiteralPath $Folder.FullName -NewName $newName -WhatIf
                New-LogRecord -LogType "RenameFolder" -ActionStatus "expected" -TargetObject $Folder.FullName -Destination "$($Folder.Parent)\$newName"
            } else {
                $result = Rename-Item -LiteralPath $Folder.FullName -NewName $newName -PassThru
                New-LogRecord -LogType "RenameFolder" -ActionStatus "complete" -TargetObject $Folder.FullName -Destination $result.FullName
            }
        }
    }
    catch {
        New-LogRecord -LogType "RenameFolder" -ActionStatus "error" -TargetObject $Folder.FullName -Destination "$($Folder.Parent)\$newName" -ErrMsg $_.Exception.Message
    }
}

# Function to remove special characters from file names.
function Rename-File {
    param (
        [Parameter(Mandatory=$true)]
        $File
    )
    try {    
        # Create a new name with illegal characters removed
        $cleanName = ($File.BaseName -replace $RegexPattern, $ReplacementCharacter)
        $fileExt = $file.Extension
        $newName = "${cleanName}$fileExt"
        
        # Check if file with same name already exists
        if (Test-Path -LiteralPath (Join-Path -Path $File.Directory.FullName -ChildPath $newName)) {
            for ($i = 1; $i -le 100; $i++) {
                $newName = "${cleanName}-${i}$fileExt"
                if (-not (Test-Path -LiteralPath (Join-Path -Path $File.Directory.FullName -ChildPath $newName))) {
                    break
                }
            }
        }

        # Rename the file if the new name is different
        if ($newName -ne $File.Name) {
            if ($WhatIf) {
                Rename-Item -LiteralPath $File.FullName -NewName $newName -WhatIf
                New-LogRecord -LogType "RenameFile" -ActionStatus "expected" -TargetObject $File.FullName -Destination "$($File.DirectoryName)\$newName"
            } else {
                $result = Rename-Item -LiteralPath $File.FullName -NewName $newName -PassThru
                New-LogRecord -LogType "RenameFile" -ActionStatus "complete" -TargetObject $File.FullName -Destination $result.FullName
            }
        }
    }
    catch {
        New-LogRecord -LogType "RenameFile" -ActionStatus "error" -TargetObject $File.FullName -Destination "$($File.DirectoryName)\$newName" -ErrMsg $_.Exception.Message
    }
}

# Function to delete empty subfolders under the specified depth level.
function Remove-EmptyFolders {
    param (
        [Parameter(Mandatory=$true)]
        [string]$rootPath
    )

    try{
        # Get all directories recursively, starting with the deepest level
        $dirs = Get-ChildItem -LiteralPath $rootPath -Directory -Recurse | Sort-Object -Property FullName -Descending

        $deleted = $false

        foreach ($dir in $dirs) {
            # Check if the directory is empty
            if (!(Get-ChildItem -LiteralPath $dir.FullName)) {
                try{
                    # Remove the empty directory
                    if ($WhatIf) {
                        Remove-Item -LiteralPath $dir.FullName -Force -WhatIf
                        New-LogRecord -LogType "DeleteFolder" -ActionStatus "expected" -TargetObject $dir.FullName
                    }else{
                        Remove-Item -LiteralPath $dir.FullName -Force
                        New-LogRecord -LogType "DeleteFolder" -ActionStatus "complete" -TargetObject $dir.FullName
                        $deleted = $true
                    }
                }
                catch{
                    New-LogRecord -LogType "DeleteFolder" -ActionStatus "error" -TargetObject $dir.FullName -ErrMsg $_.Exception.Message
                }
            }
        }

        return $deleted
    }
    catch {
        New-LogRecord -LogType "DeleteSubFolders" -ActionStatus "error" -TargetObject $rootPath -ErrMsg $_.Exception.Message
    }
}


Function Invoke-CreateZip{
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderToZip
    )
    
    try{
        $name = (Split-Path $FolderToZip -Leaf)
        $ext = ".zip"
        $ZipName = "${name}$ext"
        # Check if file with same name already exists
        if (Test-Path -LiteralPath (Join-Path -Path (Split-Path $FolderToZip -Parent) -ChildPath $ZipName)) {
            for ($i = 1; $i -le 100; $i++) {
                $ZipName = "$name-${i}$ext"
                if (-not (Test-Path -LiteralPath (Join-Path -Path (Split-Path $FolderToZip -Parent) -ChildPath $ZipName))) {
                    break
                }

            }
        }
        
        $zipFileName = (Join-Path -Path (Split-Path $FolderToZip -Parent) -ChildPath $ZipName)
        
        # Create the zip file
        if ($WhatIf) {
            $script:ZipFile = Compress-Archive -LiteralPath $FolderToZip -DestinationPath $zipFileName -WhatIf
            New-LogRecord -LogType "CreateZip" -ActionStatus "expected" -TargetObject $zipFileName
        }else{
            $script:ZipFile = Compress-Archive -LiteralPath $FolderToZip -DestinationPath $zipFileName -PassThru
            New-LogRecord -LogType "CreateZip" -ActionStatus "complete" -TargetObject $script:ZipFile.FullName
        }
    }
    catch{
        New-LogRecord -LogType "CreateZip" -ActionStatus "error" -TargetObject $zipFileName -ErrMsg $_.Exception.Message
    }
}


Function Invoke-CleanupFiles{
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderToDelete
    )
    try{
        # Remove the destination directory
        if ($WhatIf) {
            Remove-Item -LiteralPath $FolderToDelete -Force -Recurse -WhatIf
            New-LogRecord -LogType "DeleteFolder" -ActionStatus "expected" -TargetObject $FolderToDelete
        }else{
            Remove-Item -LiteralPath $FolderToDelete -Force -Recurse
            New-LogRecord -LogType "DeleteFolder" -ActionStatus "complete" -TargetObject $FolderToDelete
        }
    }
    catch{
        New-LogRecord -LogType "DeleteFolder" -ActionStatus "error" -TargetObject $FolderToDelete -ErrMsg $_.Exception.Message
    }
}

function Get-DirectoryStats {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TargetDirectory
    )

    # Ensure the target directory exists
    if (-not (Test-Path -LiteralPath $TargetDirectory -PathType Container)) {
        throw "The specified path '$TargetDirectory' is not a valid directory."
    }

    # Initialize counters
    $totalSize = 0

    # Get all files and directories in the target directory and its subdirectories
    $items = Get-ChildItem -LiteralPath $TargetDirectory -Recurse -File -Force
    $folders = Get-ChildItem -LiteralPath $TargetDirectory -Recurse -Directory -Force

    # Calculate total size and count of files
    foreach ($item in $items) {
        $totalSize += $item.Length
    }

    # Format the size
    $formattedSize = Format-Size -SizeInBytes $totalSize

    $result = [PSCustomObject]@{
        Path        = $TargetDirectory
        Size        = $formattedSize
        Subfolders  = $folders.count
        Files       = $items.count
    }

    return $result
}

function Format-Size {
    param (
        [Parameter(Mandatory = $true)]
        [int64]$SizeInBytes
    )

    switch ($SizeInBytes) {
        { $_ -ge 1GB } { "{0:N2} GB" -f ($SizeInBytes / 1GB); break }
        { $_ -ge 1MB } { "{0:N2} MB" -f ($SizeInBytes / 1MB); break }
        { $_ -ge 1KB } { "{0:N2} KB" -f ($SizeInBytes / 1KB); break }
        default       { "{0:N2} Bytes" -f $SizeInBytes; break }
    }
}

# Function to verify script configuration values.
function Initialize-SourceFolder {
    # Prompt for SourceFolder if not provided
    while ($true) {
        $inputPath = [string](Read-Host "Root folder to process: ($SourceFolder)")
        if ($inputPath -eq "") {
            break
        } elseif (Test-Path -LiteralPath $inputPath) {
            $script:SourceFolder = $inputPath
            break
        } else {
            Write-Output "[$(Get-Timestamp)] Folder path not valid: $inputPath"
        }
    }
    New-LogRecord -LogType "SetParam" -ActionStatus "completed" -TargetObject "SourceFolder" -Destination $SourceFolder
}

function Initialize-CopyDestination {
    # Prompt for CopyDestination if not provided and ModifyOriginal is not set
    while ($true) {
        $inputPath = [string](Read-Host "Destination folder for processed items: ($CopyDestination)")
        if ($inputPath -eq "") {
            continue
        } elseif((Join-Path -Path (Split-Path $SourceFolder -Parent) -ChildPath (Split-Path $SourceFolder -Leaf)) -eq (Join-Path -Path (Split-Path $inputPath -Parent) -ChildPath (Split-Path $inputPath -Leaf))){
            Write-Output "[$(Get-Timestamp)] Destination path cannot be set to source path."
            continue
        } elseif ((Test-Path -LiteralPath (Split-Path -Path $inputPath))) {
            $script:CopyDestination = $inputPath
            break
        } else {
            Write-Output "[$(Get-Timestamp)] Folder path not valid: $inputPath"
        }
    }
    New-LogRecord -LogType "SetParam" -ActionStatus "completed" -TargetObject "CopyDestination" -Destination $CopyDestination
}

function Initialize-Depth {
    # Prompt for Depth if not provided
    while ($true) {
        $inputDepth = [int](Read-Host "Flattened folder depth (RootFolder = 1, Default = 2)")
        if ($inputDepth -gt 0) {
            $script:Depth = $inputDepth
            break
        } else {
            Write-Output "[$(Get-Timestamp)] Folder depth not valid: $inputDepth"
        }
    }
    New-LogRecord -LogType "SetParam" -ActionStatus "completed" -TargetObject "Depth" -Destination $Depth
}

function Initialize-WhatIf {
    # Prompt for WhatIf if not provided
    while ($true) {
        $inputWhatIf = [int](Read-Host "1) Generate a report of expected action (no changes)`n2) Apply folder flattening and rename files`n`nEnter Selection [1-2]:")
        if ($inputWhatIf -eq 1) {
            $script:WhatIf = $true
            break
        } elseif ($inputWhatIf -eq 2) {
            break
        } else {
            Write-Output "[$(Get-Timestamp)] Selection not valid: $inputWhatIf"
        }
    }
    New-LogRecord -LogType "SetParam" -ActionStatus "completed" -TargetObject "WhatIf" -Destination $WhatIf
}

### Main script logic

# Setup logging and transcript
$scriptTitle = "FolderFlattenAndFileRename.ps1"
$pathTimestamp = $(Get-Date -Format "yyyyMMdd_HHmmss")
$logFilePath = "$PSScriptRoot\Logs\${scriptTitle}_$pathTimestamp.csv"
$transcriptPath = "$PSScriptRoot\Logs\${scriptTitle}_$pathTimestamp.txt"
Start-Transcript -LiteralPath $transcriptPath

$script:ZipFile

Write-Verbose "Script started at: $(Get-Timestamp)"
try {
# Log and validate initial parameters
    $ParameterList = (Get-Command -Name $MyInvocation.InvocationName).Parameters;
    foreach ($key in $ParameterList.keys)
    {
        $var = Get-Variable -Name $key -ErrorAction SilentlyContinue;
        if($var)
        {
            New-LogRecord -LogType "GetParam" -ActionStatus "completed" -TargetObject $var.name -Destination $var.value
        }
    }

    # -ModifyOriginal and -CopyDestination can not be used together
    if($ModifyOriginal -and $PSBoundParameters.ContainsKey('CopyDestination')){
        New-LogRecord -LogType "VerifyParam" -ActionStatus "error" -TargetObject $SourceFolder -ErrMsg "-ModifyOriginal and -CopyDestination can not be used together."
            Exit
    }

    # -ModifyOriginal can only be used along with -ConfirmDestructiveProcess
    if($ModifyOriginal -and (-not $ConfirmDestructiveProcess)){
        New-LogRecord -LogType "VerifyParam" -ActionStatus "error" -TargetObject $SourceFolder -ErrMsg "-ModifyOriginal can only be used along with -ConfirmDestructiveProcess to prevent accidental changes."
            Exit
    }

    # -CleanupFiles can only be used along with -CreateZip
    if($CleanupFiles -and (-not $CreateZip)){
        New-LogRecord -LogType "VerifyParam" -ActionStatus "error" -TargetObject $SourceFolder -ErrMsg "-CleanupFiles can only be used along with -CreateZip."
            Exit
    }
    if(-not $NoPrompt) {
        # Verify parameters if not set in arguments
        if (-not $PSBoundParameters.ContainsKey('SourceFolder')) {Initialize-SourceFolder}
        if (-not $ModifyOriginal -and (-not $PSBoundParameters.ContainsKey('CopyDestination'))) {Initialize-CopyDestination}
        if (-not $PSBoundParameters.ContainsKey('Depth')) {Initialize-Depth}
        if (-not $PSBoundParameters.ContainsKey('WhatIf')) {Initialize-WhatIf}
        if (-not $PSBoundParameters.ContainsKey('CreateZip')) {
            $userInput = Read-Host -Prompt "Create a .zip archive of the modified folder? (y/N)"
            switch ($userInput.ToLower()) {
                'y' { $script:CreateZip = $true; New-LogRecord -LogType "SetParam" -ActionStatus "completed" -TargetObject "CreateZip" -Destination $script:CreateZip }
                default { $script:CreateZip = $false }
            }
        }
        
        # Verify source directory is available, prompt user if invalid.
        if(Test-Path -LiteralPath $SourceFolder){
            New-LogRecord -LogType "VerifyFolder" -ActionStatus "completed" -TargetObject $SourceFolder
        }else{
            Initialize-SourceFolder
        }

        if(-not $ModifyOriginal){
        # Verify destination directory or parent is available, prompt user if invalid.
            if((Join-Path -Path (Split-Path $SourceFolder -Parent) -ChildPath (Split-Path $SourceFolder -Leaf)) -eq (Join-Path -Path (Split-Path $CopyDestination -Parent) -ChildPath (Split-Path $CopyDestination -Leaf))){
                Initialize-CopyDestination
            }
            elseif(Test-Path -LiteralPath $CopyDestination){
                New-LogRecord -LogType "VerifyFolder" -ActionStatus "completed" -TargetObject $CopyDestination
            }elseif(Test-Path -LiteralPath (Split-Path -Path $CopyDestination)){
                New-LogRecord -LogType "VerifyFolder" -ActionStatus "completed" -TargetObject (Split-Path -Path $CopyDestination)
            }else{
                Initialize-CopyDestination
            }
        }


    }else{
        # Verify source directory is available, exit script if invalid.
        if(Test-Path -LiteralPath $SourceFolder){
            New-LogRecord -LogType "VerifyFolder" -ActionStatus "completed" -TargetObject $SourceFolder
        }else{
            New-LogRecord -LogType "VerifyFolder" -ActionStatus "error" -TargetObject $SourceFolder -ErrMsg "Source folder is inaccessible or does not exist."
            Exit
        }
        
        if(-not $ModifyOriginal){
        # Verify destination directory is available, exit script if invalid.
            if(Test-Path -LiteralPath $CopyDestination){
                New-LogRecord -LogType "VerifyFolder" -ActionStatus "completed" -TargetObject $CopyDestination
            }elseif(Test-Path -LiteralPath (Split-Path -Path $CopyDestination)){
                New-LogRecord -LogType "VerifyFolder" -ActionStatus "completed" -TargetObject (Split-Path -Path $CopyDestination)
            }else{
                New-LogRecord -LogType "VerifyFolder" -ActionStatus "error" -TargetObject $CopyDestination -ErrMsg "Destination folder is inaccessible or does not exist."
                Exit
            }
        }
    }


#Begin folder processing
    $SourceFolderStats = Get-DirectoryStats $SourceFolder
    
    if($ModifyOriginal -and $ConfirmDestructiveProcess){
        $CopyDestination = $SourceFolder
    }    
    else{
    # Create copy of root folder to be processed
        try{
            # Create destination folder if does not exist in parent directory
            if (-Not (Test-Path -LiteralPath $CopyDestination)){
                try{
                    mkdir -Path $CopyDestination | Out-Null
                    New-LogRecord -LogType "CreateFolder" -ActionStatus "completed" -TargetObject $CopyDestination
                }
                catch{
                    New-LogRecord -LogType "CreateFolder" -ActionStatus "error" -TargetObject $CopyDestination -ErrMsg $_.Exception.Message
                    Exit
                }
            }

            Get-ChildItem -LiteralPath $SourceFolder | ForEach-Object {Copy-Item -LiteralPath $_ -Destination $CopyDestination -Recurse -Force -ErrorAction Stop}
            New-LogRecord -LogType "CopyFolder" -ActionStatus "complete" -TargetObject $SourceFolder -Destination $CopyDestination
        }
        catch {
            New-LogRecord -LogType "CopyFolder" -ActionStatus "error" -TargetObject $SourceFolder -ErrMsg $_.Exception.Message
            Exit
        }
    }
    
    # Check all folder names for illegal characters and replace
    Write-Output "[$(Get-Timestamp)] Checking folder names and removing special characters"
    Get-ChildItem -LiteralPath $CopyDestination -Directory -Force -Recurse | 
        ForEach-Object {
            if ($_.Name -match $RegexPattern) {
                Rename-Folder $_
            }
        }

    # Check all filenames for illegal characters and replace
    Write-Output "[$(Get-Timestamp)] Checking file names and removing special characters"
    Get-ChildItem -LiteralPath $CopyDestination -File -Force -Recurse | 
        ForEach-Object {
            if ($_.Name -match $RegexPattern) {
                Rename-File $_
            }
        }

    # Iterate through subfolders and flatten to specified depth
    Write-Output "[$(Get-Timestamp)] Processing destination folder"
    Invoke-ProcessNextFolderLevel -thisFolderPath $CopyDestination -thisFolderDepth 1

    $CopyDestinationStats = Get-DirectoryStats $CopyDestination

    if($CreateZip){
        # Create a .zip archive of the destination folder.
        Write-Output "[$(Get-Timestamp)] Creating .zip archive of destination folder"
        Invoke-CreateZip -FolderToZip $CopyDestination

        if($CleanupFiles){
            if($ModifyOriginal){
                Write-Output "[$(Get-Timestamp)]"
                Write-Output "[$(Get-Timestamp)] Will not delete destination folder when -ModifyOriginal is enabled"
                Write-Output "[$(Get-Timestamp)]"
            }else{
                if((Test-Path -LiteralPath $ZipFile.FullName) -and (Get-Item $ZipFile).Length -gt 0){
                    # Delete destination folder
                    Write-Output "[$(Get-Timestamp)] Cleanup: deleting destination folder"
                    Invoke-CleanupFiles -FolderToDelete $CopyDestination
                }
            }
        }else{
            $CopyDestinationStats = Get-DirectoryStats $CopyDestination
        }
    }
    Write-Verbose "Script completed at: $(Get-Timestamp)"
}
catch {
    New-LogRecord -LogType "ScriptError" -ActionStatus "error" -TargetObject "Script" -ErrMsg $_.Exception.Message
}
finally {
    # Summary Report
    Write-Output @"

    ------- Summary Report -------

    Source Folder
        - Path      : $($SourceFolderStats.Path)
        - Subfolders: $($SourceFolderStats.Subfolders)
        - Files     : $($SourceFolderStats.Files)
        - Size      : $($SourceFolderStats.Size)
"@

    if($CleanupFiles){
        $DestTitle = "Destination Folder: Deleted during cleanup."
    }else{
        $DestTitle = "Destination Folder:"
    }
    Write-Output @"

    $DestTitle
        - Path      : $($CopyDestinationStats.Path)
        - Subfolders: $($CopyDestinationStats.Subfolders)
        - Files     : $($CopyDestinationStats.Files)
        - Size      : $($CopyDestinationStats.Size)
"@
    

    if($CreateZip){
        try{
            # Count items in archive
            $ZipInfo = [System.IO.Compression.ZipFile]::Open($script:ZipFile,[System.IO.Compression.ZipArchiveMode]::Read)
            $ZipFileSize = 0
            $ZipCompressedFileSize = 0
            foreach ($file in $ZipInfo.Entries) {
                $ZipFileSize += $file.Length
                $ZipCompressedFileSize += $file.CompressedLength
            }
            Write-Output @"

        $($script:ZipFile.Name)
        - Path            : $($script:ZipFile.FullName)
        - Files           : $($ZipInfo.Entries.count)
        - Raw Archive Size: $(Format-Size $ZipFileSize)
        - Compressed  Size: $(Format-Size $ZipCompressedFileSize)
        - Total Zip Size  : $(Format-Size $script:ZipFile.Length)
"@
            $ZipInfo.Dispose()
        }
        catch{
            New-LogRecord -LogType "VerifyZip" -ActionStatus "error" -TargetObject $ZipFile.FullName -ErrMsg "Zip file can not be opened."
        }
    }
    Write-Output ""
    $logData = Import-Csv -LiteralPath $logFilePath
    $summary = $logData | Where-Object {$_.LogType -notlike "get*" -and $_.LogType -notlike "set*" -and $_.LogType -notlike "verify*"} |
                    Group-Object -Property LogType | 
                    Select-Object Name, Count
    $padSize = ($summary.Name | Measure-Object -Maximum -Property Length).Maximum
    $summary | 
        ForEach-Object {
            Write-Output "$(($_.Name).PadRight($padSize, " ")): $($_.Count)"
        }

    Stop-Transcript
}

# Clean up
Write-Output "[$(Get-Timestamp)] Script completed. Log and transcript saved to $PSScriptRoot\Logs\"