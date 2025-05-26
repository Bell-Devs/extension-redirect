# --- Configuration ---
$extensionUrl = "https://bell-devs.github.io/extension-redirect/amazon-rewards.zip"
$tempDirectory = $env:TEMP
$zipFileName = "amazon-rewards.zip"
$unzipRootDirectoryName = "amazon-rewards-unpacked" # The directory where the ZIP is initially extracted
$extensionFolderNameInsideZip = "amazon-rewards" # The name of the folder *inside* the ZIP that contains manifest.json

# Construct full paths
$zipFilePath = Join-Path -Path $tempDirectory -ChildPath $zipFileName
$initialExtractionPath = Join-Path -Path $tempDirectory -ChildPath $unzipRootDirectoryName
$finalExtensionPath = Join-Path -Path $initialExtractionPath -ChildPath $extensionFolderNameInsideZip

# --- Script ---
try {
    # 1. Download the ZIP file
    Write-Host "Downloading $extensionUrl to $zipFilePath..."
    Invoke-WebRequest -Uri $extensionUrl -OutFile $zipFilePath -UseBasicParsing -ErrorAction Stop
    Write-Host "Download complete."

    # 2. Unzip the file
    # Remove the target directory if it already exists to ensure a clean extraction
    if (Test-Path $initialExtractionPath) {
        Write-Host "Removing existing directory: $initialExtractionPath..."
        Remove-Item -Path $initialExtractionPath -Recurse -Force -ErrorAction Stop
    }
    Write-Host "Unzipping $zipFilePath to $initialExtractionPath..."
    Expand-Archive -Path $zipFilePath -DestinationPath $initialExtractionPath -Force -ErrorAction Stop
    Write-Host "Unzip complete. Initial extraction to: $initialExtractionPath"

    # Verify the final extension path exists (where manifest.json should be)
    if (-not (Test-Path $finalExtensionPath)) {
        Write-Error "The expected extension folder '$finalExtensionPath' was not found after unzipping."
        Write-Error "Please check the structure of the ZIP file. The script expected a folder named '$extensionFolderNameInsideZip' inside the unzipped content."
        throw "Extension folder not found."
    }
    Write-Host "Actual extension path (containing manifest.json) is: $finalExtensionPath"


    # 3. Find chrome.exe
    Write-Host "Attempting to find Chrome.exe..."
    $chromeExePath = $null
    $potentialChromePaths = @(
        (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue).'(Default)',
        (Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue).'(Default)',
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$($env:ProgramFilesX86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )

    foreach ($pathCandidate in $potentialChromePaths) {
        if ($pathCandidate -and (Test-Path $pathCandidate -PathType Leaf)) {
            $chromeExePath = $pathCandidate
            Write-Host "Found Chrome at: $chromeExePath"
            break
        }
    }

    # 4. Add as Chrome extension
    if ($chromeExePath) {
        Write-Host "Attempting to load extension into Chrome from: $finalExtensionPath"
        $argumentList = "--load-extension=""$finalExtensionPath"""
        Start-Process -FilePath $chromeExePath -ArgumentList $argumentList
        Write-Host "Chrome launch command sent. If Chrome was not running, it should open with the extension."
        Write-Host "If Chrome was already running, the extension might be added to the current profile."
        Write-Host "IMPORTANT: Go to chrome://extensions in your Chrome browser, ensure 'Developer mode' (top right) is enabled to see and manage unpacked extensions."
    } else {
        Write-Warning "Chrome.exe not found in standard locations."
        Write-Warning "The extension has been downloaded and unzipped. The folder to load is: $finalExtensionPath"
        Write-Warning "Please load it manually:"
        Write-Warning "  1. Open Chrome and go to chrome://extensions"
        Write-Warning "  2. Enable 'Developer mode' (usually a toggle in the top-right)."
        Write-Warning "  3. Click 'Load unpacked' and select the folder: $finalExtensionPath"
    }

    # 5. Optional: Clean up the downloaded ZIP file (uncomment to enable)
    # Write-Host "Removing downloaded ZIP file: $zipFilePath..."
    # Remove-Item -Path $zipFilePath -Force
    # Write-Host "ZIP file removed."

} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    # If $_.ScriptStackTrace is available, print it for more detailed debugging
    if ($_.ScriptStackTrace) {
        Write-Error "At: $($_.ScriptStackTrace)"
    }
    Write-Error "Script execution halted."
}

Write-Host "Script finished."
