@echo off
setlocal EnableDelayedExpansion

REM ========================================================
REM Markdown Image Organizer (UTF-8 Compatible) - Fixed
REM ========================================================

REM 1. Check if running on Desktop to prevent clutter
set "current_dir=%~dp0"
set "current_dir=%current_dir:~0,-1%"

REM Get the current user's Desktop path dynamically
for /f "usebackq tokens=2,*" %%A in (`reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop`) do (
    set "desktop_path=%%B"
)
REM Expand variable if it contains %USERPROFILE%
call set "desktop_path=%desktop_path%"

if /i "%current_dir%"=="%desktop_path%" (
    echo [ERROR] This script cannot be run directly on the Desktop root to avoid clutter.
    echo Please move the script and your markdown files to a subfolder on the Desktop or elsewhere.
    pause
    exit /b
)

REM 2. Prepare Logging Directory
set "log_dir=%current_dir%\ImageBatLog"
if not exist "%log_dir%" mkdir "%log_dir%"

REM Get current timestamp for log fileName (Format: log_YYYYMMDD_HHMMSS.txt)
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set "log_Name=log_%datetime:~0,8%_%datetime:~8,6%.txt"
set "log_file=%log_dir%\%log_Name%"

REM Initialize Log File (Creating a dummy file to ensure path exists)
echo Initializing Log... > "%log_file%"

REM 3. Execute PowerShell Script for Logic Processing
REM We pass the current directory and log file path to PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $scriptPath = '%~f0'; $currentDir = '%current_dir%'; $logFile = '%log_file%'; [System.IO.File]::ReadAllText($scriptPath) | Select-String '(?ms)^<POWERSHELL_CODE>(.*)^</POWERSHELL_CODE>' | ForEach-Object { Invoke-Expression $_.Matches.Groups[1].Value } }"

echo.
echo ========================================================
echo Processing Complete.
echo Log saved to: %log_file%
echo ========================================================
pause
goto :EOF

REM ========================================================
REM PowerShell Logic Block (Embedded)
REM ========================================================
<POWERSHELL_CODE>
# Set encoding to handle Chinese characters correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Log {
    param (
        [string]$Message,
        [string]$Path
    )
    # Append content to log file using UTF-8 encoding
    $enc = [System.Text.Encoding]::UTF8
    [System.IO.File]::AppendAllText($Path, "$Message`r`n", $enc)
    Write-Host $Message
}

$workingDir = $currentDir
$logPath = $logFile

# Clear the initial dummy text and write header
$header = @"
========================================================
MARKDOWN IMAGE MIGRATION LOG
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Working Directory: $workingDir
========================================================
"@
[System.IO.File]::WriteAllText($logPath, "$header`r`n", [System.Text.Encoding]::UTF8)

# Get all .md files in the current directory (Top-level only)
$mdFiles = Get-ChildItem -Path $workingDir -Filter "*.md" -File

if ($mdFiles.Count -eq 0) {
    Write-Log "No Markdown files found in the directory." $logPath
    exit
}

foreach ($file in $mdFiles) {
    Write-Log "`n[PROCESSING FILE]: $($file.Name)" $logPath
    
    $content = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    $fileNameNoExt = $file.BaseName
    $imageFolder = Join-Path -Path $workingDir -ChildPath "${fileNameNoExt}Image"
    
    # Regex to find images: ![alt](path)
    $regex = '!\[(.*?)\]\((.*?)\)'
    
    $matches = [regex]::Matches($content, $regex)
    
    if ($matches.Count -eq 0) {
        Write-Log "  - No images found." $logPath
        continue
    }

    # Create image folder if it doesn't exist (only if images are found)
    if (-not (Test-Path -LiteralPath $imageFolder)) {
        New-Item -Path $imageFolder -ItemType Directory | Out-Null
        Write-Log "  - Created directory: $imageFolder" $logPath
    }

    # Normalize image folder full path for comparisons
    try {
        $imageFolderFull = (Get-Item -LiteralPath $imageFolder -ErrorAction Stop).FullName
    } catch {
        # Fallback: combine and normalize
        $imageFolderFull = [System.IO.Path]::GetFullPath($imageFolder)
    }

    $newContent = $content
    $changesMade = $false

    foreach ($match in $matches) {
        $originalString = $match.Value
        $altText = $match.Groups[1].Value
        $imagePath = $match.Groups[2].Value

        # Determine source path (absolute or relative)
        if ($imagePath -match '^[a-zA-Z]:\\' -or $imagePath -match '^\\\\') {
            $sourceCandidate = $imagePath
        } else {
            # If path starts with ./ or .\, remove that prefix when joining to avoid doubled './'
            if ($imagePath -match '^[.][\\/](.*)') {
                $relPart = $Matches[1]
                $sourceCandidate = Join-Path -Path $workingDir -ChildPath $relPart
            } else {
                $sourceCandidate = Join-Path -Path $workingDir -ChildPath $imagePath
            }
        }

        # Try to get full normalized source path if the file exists
        $sourceFull = $null
        try {
            if (Test-Path -LiteralPath $sourceCandidate -PathType Leaf) {
                $sourceFull = (Get-Item -LiteralPath $sourceCandidate -ErrorAction Stop).FullName
            } else {
                # If it doesn't exist, also try to interpret raw imagePath as absolute (in case imagePath already was like ./NameImage/xxx.png)
                if (Test-Path -LiteralPath $imagePath -PathType Leaf) {
                    $sourceFull = (Get-Item -LiteralPath $imagePath -ErrorAction Stop).FullName
                }
            }
        } catch {}

        if (-not $sourceFull) {
            Write-Log "  [FAILURE] Source image not found: $imagePath" $logPath
            continue
        }

        # Compute destination full path
        $imageFileName = Split-Path -Path $sourceFull -Leaf
        $destPathCandidate = Join-Path -Path $imageFolderFull -ChildPath $imageFileName
        $destFull = [System.IO.Path]::GetFullPath($destPathCandidate)

        # If source and dest are identical, skip copying
        if ($sourceFull -ieq $destFull) {
            Write-Log "  [SKIP] Source and destination are the same; no copy needed: $sourceFull" $logPath
            # Also ensure the markdown link is in desired relative form; if it already points to same, no change
            # If the markdown still references an absolute path and you want to normalize it, update it here:
            $expectedRel = "./${fileNameNoExt}Image/$imageFileName"
            if ($originalString -notmatch [regex]::Escape($expectedRel)) {
                $newContent = $newContent.Replace($originalString, "![$altText]($expectedRel)")
                $changesMade = $true
                Write-Log "    - Updated markdown link to relative path: $expectedRel" $logPath
            }
            continue
        }

        # Attempt to copy; only log success if copy actually completed
        try {
            Copy-Item -LiteralPath $sourceFull -Destination $destFull -Force -ErrorAction Stop

            if (Test-Path -LiteralPath $destFull -PathType Leaf) {
                $relPath = "./${fileNameNoExt}Image/$imageFileName"
                $newContent = $newContent.Replace($originalString, "![$altText]($relPath)")
                Write-Log "  [SUCCESS] Copied: $sourceFull -> $destFull" $logPath
                $changesMade = $true
            } else {
                Write-Log "  [FAILURE] Copy reported no error but destination not found: $sourceFull" $logPath
            }
        } catch {
            Write-Log "  [ERROR] Copy failed: $sourceFull -> $destFull. Exception: $($_.Exception.Message)" $logPath
        }
    }

    if ($changesMade) {
        try {
            [System.IO.File]::WriteAllText($file.FullName, $newContent, [System.Text.Encoding]::UTF8)
            Write-Log "  - Updated Markdown file: $($file.Name)" $logPath
        } catch {
            Write-Log "  - Failed to write updated Markdown content." $logPath
        }
    } else {
        Write-Log "  - No changes required for this file." $logPath
    }
}
</POWERSHELL_CODE>