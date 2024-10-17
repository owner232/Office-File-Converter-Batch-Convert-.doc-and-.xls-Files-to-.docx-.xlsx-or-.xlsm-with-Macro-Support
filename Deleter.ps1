# use this to find already converted xlsx/xlsm and docx files, Then delete the old .doc or .xls file.
# Set the path to the folder where your .doc and .xls files are located
$folderPath = "c:/test/test/test"  # <-- Update this path

# Check if the folder exists
if (-not (Test-Path -Path $folderPath)) {
    Write-Host "The folder path '$folderPath' does not exist. Please check the path."
    exit
}

# Arrays to store files without counterparts
$filesWithoutCounterpart = @()
$filesToDelete = @()

# Function to log actions
function Log-Action {
    param (
        [string]$message,
        [string]$logFilePath
    )
    Add-Content -Path $logFilePath -Value "$message"
}

# Function to check for counterpart files
function Check-Counterpart {
    param (
        [string]$filePath,
        [string]$extensionToCheck1,
        [string]$extensionToCheck2
    )
    
    $counterpart1 = [System.IO.Path]::ChangeExtension($filePath, $extensionToCheck1)
    $counterpart2 = [System.IO.Path]::ChangeExtension($filePath, $extensionToCheck2)

    if (-not (Test-Path $counterpart1) -and -not (Test-Path $counterpart2)) {
        return $false
    }
    return $true
}

# Scanning for .doc and .xls files
Write-Host "Scanning directory: $folderPath"
$docFiles = Get-ChildItem -Path $folderPath -Recurse -Filter *.doc | Where-Object { $_.Extension -ieq ".doc" -and $_.Name -notlike "*.docx" }
$xlsFiles = Get-ChildItem -Path $folderPath -Recurse -Filter *.xls | Where-Object { $_.Extension -ieq ".xls" -and $_.Name -notlike "*.xlsx" -and $_.Name -notlike "*.xlsm" }

# Checking for .doc files that do not have a .docx counterpart
foreach ($docFile in $docFiles) {
    $docFilePath = $docFile.FullName
    if (-not (Check-Counterpart -filePath $docFilePath -extensionToCheck1 ".docx" -extensionToCheck2 "")) {
        Write-Host "No .docx counterpart for: $docFilePath"
        $filesWithoutCounterpart += $docFilePath
    } else {
        $filesToDelete += $docFilePath
    }
}

# Checking for .xls files that do not have a .xlsx or .xlsm counterpart
foreach ($xlsFile in $xlsFiles) {
    $xlsFilePath = $xlsFile.FullName
    if (-not (Check-Counterpart -filePath $xlsFilePath -extensionToCheck1 ".xlsx" -extensionToCheck2 ".xlsm")) {
        Write-Host "No .xlsx or .xlsm counterpart for: $xlsFilePath"
        $filesWithoutCounterpart += $xlsFilePath
    } else {
        $filesToDelete += $xlsFilePath
    }
}

# Displaying files without counterparts
if ($filesWithoutCounterpart.Count -gt 0) {
    Write-Host "`nThe following .doc and .xls files do not have corresponding .docx, .xlsx, or .xlsm files:"
    $filesWithoutCounterpart | ForEach-Object { Write-Host $_ }
} else {
    Write-Host "All .doc and .xls files have corresponding .docx, .xlsx, or .xlsm counterparts."
}

# Asking if the user wants to delete the .doc and .xls files that have counterparts
if ($filesToDelete.Count -gt 0) {
    $deleteFilesResponse = Read-Host "`nWould you like to delete the original .doc and .xls files that have corresponding .docx/.xlsx/.xlsm files? (yes/no)"
    if ($deleteFilesResponse -eq 'yes') {
        foreach ($fileToDelete in $filesToDelete) {
            try {
                Remove-Item -Path $fileToDelete -Force
                Write-Host "Deleted: $fileToDelete"
            } catch {
                Write-Host "ERROR: Failed to delete $fileToDelete"
            }
        }
    } else {
        Write-Host "No files were deleted."
    }
} else {
    Write-Host "No files were found to delete."
}

Write-Host "`nOperation completed."
