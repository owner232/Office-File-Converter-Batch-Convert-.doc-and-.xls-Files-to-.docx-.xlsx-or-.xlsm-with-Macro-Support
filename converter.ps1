# Set the path to the folder where your .doc and .xls files are located
$folderPath = "c:/test/test/test"  # <--- Update this path

# Check if folder exists
if (-not (Test-Path -Path $folderPath)) {
    Write-Host "The folder path '$folderPath' does not exist. Please check the path."
    exit
}

# Create or append to the conversion log and delete log files
$logFilePath = Join-Path $folderPath "conversion_log.txt"
$deleteLogFilePath = Join-Path $folderPath "delete_log.txt"
Add-Content -Path $logFilePath -Value "Conversion Log - $(Get-Date)" -Force
Add-Content -Path $deleteLogFilePath -Value "Deletion Log - $(Get-Date)" -Force

# Function to log actions
function Log-Action {
    param (
        [string]$message,
        [string]$logFilePath
    )
    Add-Content -Path $logFilePath -Value "$message"
}

# Ask user if they want to convert .doc files
$convertDocs = $false
$convertDocsResponse = Read-Host "Do you want to scan and convert the .doc files to .docx? (yes/no)"
if ($convertDocsResponse -eq 'yes') {
    $convertDocs = $true
}

# Ask user if they want to convert .xls files
$convertXls = $false
$convertXlsResponse = Read-Host "Do you want to scan and convert the .xls files to .xlsx or .xlsm (for macro-enabled)? (yes/no)"
if ($convertXlsResponse -eq 'yes') {
    $convertXls = $true
}

# If no conversion is selected, exit the script
if (-not $convertDocs -and -not $convertXls) {
    Write-Host "No conversions selected. Exiting."
    exit
}

# Scanning for files
Write-Host "I am scanning the directory: $folderPath"
Log-Action "Scanning directory: $folderPath" $logFilePath

# Arrays for files to convert and files to prompt for reconversion
$filesToConvertDoc = @()
$filesToConvertXls = @()
$filesToReconvertDoc = @()
$filesToReconvertXls = @()

# Process .doc files if selected
if ($convertDocs) {
    # Collect all .doc files, ignoring those that are .docx
    $docFiles = Get-ChildItem -Path $folderPath -Recurse -Filter *.doc | Where-Object {
        $_.Extension -ieq ".doc" -and $_.Name -notlike "*.docx"
    }

    foreach ($docFile in $docFiles) {
        $docFilePath = $docFile.FullName
        $docxFilePath = [System.IO.Path]::ChangeExtension($docFilePath, ".docx")

        # Check if both the .doc and .docx files exist in the same directory
        if (Test-Path $docxFilePath) {
            $filesToReconvertDoc += $docFilePath
            Log-Action "Detected existing .docx in the same directory: $docFilePath has already been converted to .docx." $logFilePath
        } else {
            $filesToConvertDoc += $docFilePath
        }
    }
}

# Process .xls files if selected
if ($convertXls) {
    # Collect all .xls files, ignoring those that are already .xlsx or .xlsm
    $xlsFiles = Get-ChildItem -Path $folderPath -Recurse -Filter *.xls | Where-Object {
        $_.Extension -ieq ".xls" -and $_.Name -notlike "*.xlsx" -and $_.Name -notlike "*.xlsm"
    }

    foreach ($xlsFile in $xlsFiles) {
        $xlsFilePath = $xlsFile.FullName
        $xlsxFilePath = [System.IO.Path]::ChangeExtension($xlsFilePath, ".xlsx")
        $xlsmFilePath = [System.IO.Path]::ChangeExtension($xlsFilePath, ".xlsm")

        # Check if both the .xls and .xlsx/.xlsm files exist in the same directory
        if ((Test-Path $xlsxFilePath) -or (Test-Path $xlsmFilePath)) {
            $filesToReconvertXls += $xlsFilePath
            Log-Action "Detected existing .xlsx/.xlsm in the same directory: $xlsFilePath has already been converted." $logFilePath
        } else {
            $filesToConvertXls += $xlsFilePath
        }
    }
}

# Output files that have already been converted and ask for reconversion
if ($filesToReconvertDoc.Count -gt 0 -or $filesToReconvertXls.Count -gt 0) {
    Write-Host "`nThe following files appear to already have been converted. Would you still like to reconvert them?"
    if ($filesToReconvertDoc.Count -gt 0) {
        Write-Host "`n.doc files that have already been converted (same name and directory):"
        $filesToReconvertDoc | ForEach-Object { Write-Host $_ }
    }
    if ($filesToReconvertXls.Count -gt 0) {
        Write-Host "`n.xls files that have already been converted (same name and directory):"
        $filesToReconvertXls | ForEach-Object { Write-Host $_ }
    }
    $reconvertResponse = Read-Host "`nWould you like to reconvert the files listed above? (yes/no)"
    if ($reconvertResponse -eq 'yes') {
        $filesToConvertDoc += $filesToReconvertDoc
        $filesToConvertXls += $filesToReconvertXls
        Log-Action "User chose to reconvert files that were already converted." $logFilePath
    } else {
        Log-Action "User chose NOT to reconvert files that were already converted." $logFilePath
    }
}

# Output files to convert and log them
if ($convertDocs -and $filesToConvertDoc.Count -gt 0) {
    Write-Host "`nFiles that will be converted from .doc to .docx:"
    $filesToConvertDoc | ForEach-Object { Write-Host $_ }
    Log-Action "Files to convert from .doc to .docx:" $logFilePath
    $filesToConvertDoc | ForEach-Object { Log-Action $_ $logFilePath }
}

if ($convertXls -and $filesToConvertXls.Count -gt 0) {
    Write-Host "`nFiles that will be converted from .xls to .xlsx or .xlsm (for macro-enabled files):"
    $filesToConvertXls | ForEach-Object { Write-Host $_ }
    Log-Action "Files to convert from .xls to .xlsx or .xlsm:" $logFilePath
    $filesToConvertXls | ForEach-Object { Log-Action $_ $logFilePath }
}

# Final confirmation before conversion
$finalConfirm = Read-Host "`nDo you want to proceed with the conversion of the listed files? (yes/no)"
if ($finalConfirm -ne 'yes') {
    Write-Host "Conversion cancelled. Exiting."
    exit
}

# Perform the conversions

# Convert .doc to .docx
if ($convertDocs -and $filesToConvertDoc.Count -gt 0) {
    Write-Host "`nStarting .doc to .docx conversion..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # Suppress alerts

    $processedDocFiles = 0
    $totalDocFiles = $filesToConvertDoc.Count
    $startTime = Get-Date
    $totalTimeDoc = [System.TimeSpan]::Zero

    foreach ($docFilePath in $filesToConvertDoc) {
        $docxFilePath = [System.IO.Path]::ChangeExtension($docFilePath, ".docx")
        try {
            $timeTaken = Measure-Command {
                $document = $word.Documents.Open($docFilePath)
                $document.SaveAs([ref] $docxFilePath, [ref] 16)  # 16 is the DOCX format code
                $document.Close()
            }

            # Update timing and progress
            $totalTimeDoc += $timeTaken
            $processedDocFiles++
            $avgTimePerFileDoc = $totalTimeDoc.TotalSeconds / $processedDocFiles
            $remainingFilesDoc = $totalDocFiles - $processedDocFiles
            $estimatedTimeRemainingDoc = $avgTimePerFileDoc * $remainingFilesDoc / 60

            Write-Host ("Converted {0} of {1} .doc files. Estimated time remaining: {2:N2} minutes" -f $processedDocFiles, $totalDocFiles, $estimatedTimeRemainingDoc)
            Log-Action ("Converted {0} of {1} .doc files. Estimated time remaining: {2:N2} minutes" -f $processedDocFiles, $totalDocFiles, $estimatedTimeRemainingDoc) $logFilePath
        }
        catch {
            Write-Host "ERROR: Failed to convert $docFilePath"
            Log-Action "ERROR: Failed to convert $docFilePath - $_" $logFilePath
        }
    }
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
}

# Convert .xls to .xlsx or .xlsm depending on if it contains macros
if ($convertXls -and $filesToConvertXls.Count -gt 0) {
    Write-Host "`nStarting .xls to .xlsx/.xlsm conversion..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $processedXlsFiles = 0
    $totalXlsFiles = $filesToConvertXls.Count
    $totalTimeXls = [System.TimeSpan]::Zero

    foreach ($xlsFilePath in $filesToConvertXls) {
        $xlsxFilePath = [System.IO.Path]::ChangeExtension($xlsFilePath, ".xlsx")
        $xlsmFilePath = [System.IO.Path]::ChangeExtension($xlsFilePath, ".xlsm")
        try {
            $workbook = $excel.Workbooks.Open($xlsFilePath)

            # Check if the workbook contains macros (VBProject)
            if ($workbook.HasVBProject -eq $true) {
                # Save as .xlsm if macros are found
                $workbook.SaveAs($xlsmFilePath, 52)  # 52 is the XLSM format code
                Log-Action "Converted $xlsFilePath to $xlsmFilePath (macro-enabled)" $logFilePath
            } else {
                # Save as .xlsx if no macros are found
                $workbook.SaveAs($xlsxFilePath, 51)  # 51 is the XLSX format code
                Log-Action "Converted $xlsFilePath to $xlsxFilePath (macro-free)" $logFilePath
            }
            $workbook.Close()

            $processedXlsFiles++
            $timeTaken = [timespan]::FromTicks(($processedXlsFiles * [TimeSpan]::FromSeconds(1).Ticks))
            $estimatedTimeRemaining = [timespan]::FromTicks(($totalXlsFiles - $processedXlsFiles) * $timeTaken.Ticks)

            Write-Host ("Converted {0} of {1} .xls files. Estimated time remaining: {2:N2} minutes" -f $processedXlsFiles, $totalXlsFiles, $estimatedTimeRemaining.TotalMinutes)
        }
        catch {
            Write-Host "ERROR: Failed to convert $xlsFilePath"
            Log-Action "ERROR: Failed to convert $xlsFilePath - $_" $logFilePath
        }
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}

# Final confirmation to delete old .doc and .xls files
$deleteFilesResponse = Read-Host "`nWould you like to delete the original .doc and .xls files? (yes/no)"
if ($deleteFilesResponse -eq 'yes') {
    # Delete original .doc files
    if ($convertDocs -and $filesToConvertDoc.Count -gt 0) {
        foreach ($docFile in $filesToConvertDoc) {
            try {
                Remove-Item -Path $docFile -Force
                Write-Host "Deleted: $docFile"
                Log-Action "Deleted: $docFile (Matching .docx exists)" $deleteLogFilePath
            }
            catch {
                Write-Host "ERROR: Failed to delete $docFile"
                Log-Action "ERROR: Failed to delete $docFile - $_" $deleteLogFilePath
            }
        }
    }

    # Delete original .xls files
    if ($convertXls -and $filesToConvertXls.Count -gt 0) {
        foreach ($xlsFile in $filesToConvertXls) {
            try {
                Remove-Item -Path $xlsFile -Force
                Write-Host "Deleted: $xlsFile"
                Log-Action "Deleted: $xlsFile (Matching .xlsx/.xlsm exists)" $deleteLogFilePath
            }
            catch {
                Write-Host "ERROR: Failed to delete $xlsFile"
                Log-Action "ERROR: Failed to delete $xlsFile - $_" $deleteLogFilePath
            }
        }
    }
}

Write-Host "`nConversion process completed!"
Log-Action "Conversion process completed." $logFilePath
