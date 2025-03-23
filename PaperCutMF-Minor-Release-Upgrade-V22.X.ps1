# Variables
$pcmfBackupLocation = $env:BackupPCMFDB_Location
$logFilePath = "$($env:HSMIT_LogFolderLocation)\papercut-upgrade.log"
$pcmfCDNUrl = "http://cdn.hsho.me/pcmf/"
$versionToInstall = $env:PCMF_TargetVersion
$pcmfInstallerName = "$($versionToInstall).exe"
$pcmfOldVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'Version' | Select-Object -ExpandProperty 'Version'
$pcmfInstallPath = Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'InstallPath' | Select-Object -ExpandProperty 'InstallPath'
$backupFolder = Join-Path -Path "$($pcmfBackupLocation)\PaperCut-Backup\$($pcmfOldVersion)" -ChildPath "PaperCut MF"

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

# Create directories if they are not present in C:\
if (-not (Test-Path "C:\HSMIT")) {
    New-Item -ItemType Directory -Path "C:\HSMIT"
    Write-LogEntry -Data "Created C:\HSMIT."
}

if (-not (Test-Path "C:\HSMIT\PaperCut-Upgrade-Files")) {
    New-Item -ItemType Directory -Path "C:\HSMIT\PaperCut-Upgrade-Files"
    Write-LogEntry -Data "Created C:\HSMIT\PaperCut-Upgrade-Files."
}

if (-not (Test-Path "C:\HSMIT\PaperCut-Upgrade-Logs")) {
    New-Item -ItemType Directory -Path "C:\HSMIT\PaperCut-Upgrade-Logs"
    Write-LogEntry -Data "Created C:\HSMIT\PaperCut-Upgrade-Logs."
}

if (-not (Test-Path "C:\HSMIT\PaperCut-Backup")) {
    New-Item -ItemType Directory -Path "C:\HSMIT\PaperCut-Backup"
    Write-LogEntry -Data "Created C:\HSMIT\PaperCut-Backup."
}

# Clear Old Backups
Get-ChildItem -Path "C:\HSMIT\PaperCut-Backup" -Recurse | Remove-Item -Force -Recurse

# Stop PaperCut Server
Write-LogEntry -Data "Stopping PaperCut Server"
Stop-Service -Name "PaperCut Application Server"

# Create Backup Folder
if (!(Test-Path -Path $backupFolder -PathType Container)) {
    Write-LogEntry -Data "Creating Backup Folder: $backupFolder"
    New-Item -Path $backupFolder -ItemType Directory
} else {
    Write-LogEntry -Data "Backup folder already exists. Skipping backup creation: $backupFolder"
}

# Copy the installation folder to the backup folder
Write-LogEntry -Data "Copying installation folder to backup: $pcmfInstallPath to $backupFolder"

#Copy-Item -Path $pcmfInstallPath -Destination $backupFolder -Recurse -Force

Compress-Archive -Path $pcmfInstallPath -DestinationPath $backupFolder -CompressionLevel Fastest

Write-LogEntry -Data "Backup complete and saved to: $backupFolder"

# Download the file
Write-LogEntry -Data "Downloading the upgrade file from $($pcmfCDNUrl)$($pcmfInstallerName) to C:\HSMIT\PaperCut-Upgrade-Files\$($pcmfInstallerName)"
try {
    Start-BitsTransfer -Source "$($pcmfCDNUrl)$($pcmfInstallerName)" -Destination "C:\HSMIT\PaperCut-Upgrade-Files\$($pcmfInstallerName)"
}
catch {
    Write-LogEntry -Data "Download Failed."
    exit
}
finally {
    # Check if the file was downloaded successfully
    if (Test-Path -Path "C:\HSMIT\PaperCut-Upgrade-Files\$($pcmfInstallerName)" -PathType Leaf) {
        # Run PaperCut Upgrade
        Write-LogEntry -Data "Running PaperCut Upgrade"
        Start-Process -FilePath "C:\HSMIT\PaperCut-Upgrade-Files\$($pcmfInstallerName)" -ArgumentList "/VERYSILENT"
        Write-LogEntry -Data "Beginning sleep period"
        Start-Sleep -Seconds $env:SleepTime
    } else {
        Write-LogEntry -Data "Error: Failed to download the upgrade file."
    }
}

# Start PaperCut Server
Write-LogEntry -Data "Starting PaperCut Server, please wait."
Start-Service -Name "PaperCut Application Server"

# Get the current version number from registry
$currentVersion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'Version' | Select-Object -ExpandProperty 'Version')
Write-LogEntry -Data "Current version: $($currentVersion)"

# Upgrade completed
Write-LogEntry -Data "Upgrade completed."