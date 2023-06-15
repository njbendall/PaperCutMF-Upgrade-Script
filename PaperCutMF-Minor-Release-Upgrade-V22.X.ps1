# Variables
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$pcmfBackupLocation = "C:\EDUIT"
$logFilePath = Join-Path -Path $scriptPath -ChildPath "papercut-upgrade.log"
$pcmfCDNUrl = "https://cdn.papercut.com/web/products/ng-mf/installers/mf/22.x/pcmf-setup-22.1.1.66714.exe"
$pcmfInstallerName = [regex]::Match($pcmfCDNUrl, '/([^/]+)$').Groups[1].Value
$pcmfOldVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'Version' | Select-Object -ExpandProperty 'Version'
$pcmfInstallPath = Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'InstallPath' | Select-Object -ExpandProperty 'InstallPath'

# Logging and Console Printing 
function Write-LogEntry {
    param(
        [string]$Data
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[${timestamp}] $data"
    Write-Output -InputObject "$($data)"
    $logEntry | Out-File -FilePath $logFilePath -Append
}

# Stop PaperCut Server
Write-LogEntry -Data "Stopping PaperCut Server"
Stop-Service -Name "PaperCut Application Server"

# Create Backup Folder
$backupFolder = Join-Path -Path "$($pcmfBackupLocation)\PaperCut-Backup\$($pcmfOldVersion)" -ChildPath "PaperCut MF"
if (!(Test-Path -Path $backupFolder -PathType Container)) {
    Write-LogEntry -Data "Creating Backup Folder: $backupFolder"
    New-Item -Path $backupFolder -ItemType Directory
} else {
    Write-LogEntry -Data "Backup folder already exists. Skipping backup creation: $backupFolder"
}

# Copy the installation folder to the backup folder
Write-LogEntry -Data "Copying installation folder to backup: $pcmfInstallPath to $backupFolder"
Copy-Item -Path $pcmfInstallPath -Destination $backupFolder -Recurse -Force

# Backup Complete
Write-LogEntry -Data "Backup complete and saved to: $backupFolder"

# Download the file
Write-LogEntry -Data "Downloading the upgrade file"
Start-BitsTransfer -Source $pcmfCDNUrl -Destination (Join-Path -Path $scriptPath -ChildPath $pcmfInstallerName)

# Check if the file was downloaded successfully
if (Test-Path -Path (Join-Path -Path $scriptPath -ChildPath $pcmfInstallerName) -PathType Leaf) {
    # Run PaperCut Upgrade
    Write-LogEntry -Data "Running PaperCut Upgrade"
    Start-Process -FilePath (Join-Path -Path $scriptPath -ChildPath $pcmfInstallerName) -ArgumentList "/VERYSILENT"
} else {
    Write-LogEntry -Data "Error: Failed to download the upgrade file."
}

# Start PaperCut Server
Write-LogEntry -Data "Starting PaperCut Server, please wait."
Start-Service -Name "PaperCut Application Server"

# Get the current version number from registry
$currentVersion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'Version' | Select-Object -ExpandProperty 'Version')
Write-LogEntry -Data "Current version: $($currentVersion)"

# Upgrade completed
Write-LogEntry -Data "Upgrade completed."
