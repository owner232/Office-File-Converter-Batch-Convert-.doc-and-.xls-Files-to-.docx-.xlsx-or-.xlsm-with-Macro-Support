Office File Converter: Batch Convert .doc and .xls Files to .docx, .xlsx, or .xlsm (with Macro Support)
Description:

This PowerShell script allows you to batch convert legacy Microsoft Office files (.doc and .xls) to their modern counterparts (.docx, .xlsx, or .xlsm for macro-enabled files). It automatically detects whether an .xls file contains macros and converts it to .xlsm to preserve those macros. If no macros are detected, it will convert to .xlsx.

The script is designed to scan directories recursively for .doc and .xls files and ensure that each file has a corresponding modern format. If a conversion counterpart already exists (.docx, .xlsx, or .xlsm), the script prompts whether you want to reconvert the file or skip it. Additionally, the script logs all actions, such as conversions and errors, for easy tracking.
Key Features:

    Batch Conversion of Legacy Office Files: Automatically convert all .doc and .xls files to .docx, .xlsx, or .xlsm (macro-enabled).
    Macro Detection: The script intelligently detects if .xls files contain macros and saves them as .xlsm files to retain macro functionality.
    Skip Re-conversion of Existing Files: If modern versions of the .doc or .xls files already exist in the same directory (.docx, .xlsx, or .xlsm), the script asks whether to skip or re-convert them.
    Recursion Through Directories: The script searches all subdirectories within a specified folder, so you can batch convert files across large, nested folder structures.
    Deletion Prompt (Optional): After converting files, the script can prompt you to delete the original .doc and .xls files to help with cleanup.

How It Works:

    Scan Folder for Files: The script scans a specified folder (and all its subfolders) for .doc and .xls files.
    Check for Modern Equivalents: The script checks if a .docx, .xlsx, or .xlsm file with the same name exists in the same folder.
    Convert Files:
        .doc files are converted to .docx.
        .xls files are converted to .xlsx if no macros are detected.
        .xls files are converted to .xlsm if macros are present.
    Prompt for Re-conversion: If modern equivalents exist, the script prompts you to either re-convert or skip those files.
    Detailed Logs: Conversion actions are logged for easy reference.

Installation:

    Download the Script: Download the PowerShell script file from this repository.
    Set Execution Policy: You may need to set your PowerShell execution policy to allow script execution:

    powershell

Set-ExecutionPolicy RemoteSigned -Scope Process

Run the Script: Open PowerShell, navigate to the folder where you saved the script, and run it:

powershell

    .\office-file-converter.ps1

Usage:

    Set the Target Folder: Edit the script to set the $folderPath variable to the folder you want to scan for legacy Office files.

    Run the Script: The script will:
        Scan the directory for .doc and .xls files.
        Check for corresponding .docx, .xlsx, or .xlsm files.
        Convert legacy files if no modern equivalent exists.
        Prompt for re-conversion if modern equivalents exist.

    Deletion Option (Optional): After conversion, the script asks if you'd like to delete the original .doc and .xls files.

Example:

powershell

Do you want to scan and convert the .doc files to .docx? (yes/no): yes
Do you want to scan and convert the .xls files to .xlsx or .xlsm (for macro-enabled)? (yes/no): yes

The following files appear to already have been converted. Would you still like to reconvert them?
.doc files that have already been converted (same name and directory):
C:\Path\To\Your\Folder\file1.doc

.xls files that have already been converted (same name and directory):
C:\Path\To\Your\Folder\file1.xls

Would you like to reconvert the files listed above? (yes/no): no

Files that will be converted from .doc to .docx:
C:\Path\To\Your\Folder\file2.doc

Files that will be converted from .xls to .xlsx or .xlsm (for macro-enabled files):
C:\Path\To\Your\Folder\file2.xls

Do you want to proceed with the conversion of the listed files? (yes/no): yes

Would you like to delete the original .doc and .xls files? (yes/no): yes

Requirements:

    PowerShell 5.0 or later
    Microsoft Office: Ensure that Microsoft Word and Excel are installed on the machine where the script is executed.

License:

This project is licensed under the MIT License. Feel free to use and modify the script as needed.
