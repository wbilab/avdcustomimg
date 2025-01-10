# Set download directory
$DownloadDirectory = "C:\Temp"
$CurrentDate = Get-Date -Format "dd-MM-yyyy"
$LogPath = "$DownloadDirectory\MicrosoftTeamsNEW-install_$CurrentDate.log"

# Starte Transkriptprotokollierung
Start-Transcript -Path $LogPath -Force

# Check if ms-teams.exe is running and terminate all its processes
$TeamsProcesses = Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue
if ($TeamsProcesses) {
    Write-Host "Microsoft Teams is running. Terminating all Microsoft Teams processes..." -ForegroundColor Cyan
    $TeamsProcesses | ForEach-Object {
        $_.Kill()
        Write-Host "Terminated process: $($_.Name) (PID: $($_.Id))" -ForegroundColor Yellow
    }
} else {
    Write-Host "Microsoft Teams is not running." -ForegroundColor Green
}


# Create download directory if it doesn't exist
if (-not (Test-Path -Path $DownloadDirectory -PathType Container)) {
    New-Item -Path $DownloadDirectory -ItemType Directory | Out-Null
}

# Function to download a file
function DownloadFile {
    param (
        [string]$Url,
        [string]$FileName
    )
    Write-Host "Downloading file: $($FileName)" -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $Url -OutFile "$DownloadDirectory\$FileName" -ErrorAction Stop
        Write-Host "Downloaded: $($FileName)" -ForegroundColor Green
    } catch {
        Write-Host "Error downloading: $($FileName)" -ForegroundColor Red
        Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

# Download latest WebRTC MSI from Microsoft
$WebRTCUrl = "https://aka.ms/msrdcwebrtcsvc/msi"
$WebRTCFileName = "MsRdcWebRTCSvc_x64.msi"
DownloadFile -Url $WebRTCUrl -FileName $WebRTCFileName

# Download latest Teams Bootstrapper and MSIX
$TeamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
$TeamsBootstrapperFileName = "teamsbootstrapper.exe"
DownloadFile -Url $TeamsBootstrapperUrl -FileName $TeamsBootstrapperFileName

$TeamsMSIXUrl = "https://go.microsoft.com/fwlink/?linkid=2196106"
$TeamsMSIXFileName = "MSTeams-x64.msix"
DownloadFile -Url $TeamsMSIXUrl -FileName $TeamsMSIXFileName


###########################################################
# Teams Classic cleanup
###########################################################

# Function to uninstall Teams Classic
function Uninstall-TeamsClassic($TeamsPath) {
    try {
        $process = Start-Process -FilePath "$TeamsPath\Update.exe" -ArgumentList "--uninstall /s" -PassThru -Wait -ErrorAction STOP

        if ($process.ExitCode -ne 0) {
            Write-Error "Uninstallation failed with exit code $($process.ExitCode)."
        }
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

# Remove Teams Machine-Wide Installer
Write-Host "Removing Classic Teams Machine-wide Installer"

#Windows Uninstaller Registry Path
$registryPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

# Get all subkeys and match the subkey that contains "Teams Machine-Wide Installer" DisplayName.
$MachineWide = Get-ItemProperty -Path $registryPath | Where-Object -Property DisplayName -eq "Teams Machine-Wide Installer"

if ($MachineWide) {
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/x ""$($MachineWide.PSChildName)"" /qn" -NoNewWindow -Wait
}
else {
    Write-Host "Classic Teams Machine-Wide Installer not found"
}

# Get all Users
$AllUsers = Get-ChildItem -Path "$($ENV:SystemDrive)\Users"

# Process all Users
foreach ($User in $AllUsers) {
    Write-Host "Processing user: $($User.Name)"

    # Locate installation folder
    $localAppData = "$($ENV:SystemDrive)\Users\$($User.Name)\AppData\Local\Microsoft\Teams"
    $programData = "$($env:ProgramData)\$($User.Name)\Microsoft\Teams"

    if (Test-Path "$localAppData\Current\Teams.exe") {
        Write-Host "  Uninstall Classic Teams for user $($User.Name)"
        Uninstall-TeamsClassic -TeamsPath $localAppData
    }
    elseif (Test-Path "$programData\Current\Teams.exe") {
        Write-Host "  Uninstall Classic Teams for user $($User.Name)"
        Uninstall-TeamsClassic -TeamsPath $programData
    }
    else {
        Write-Host "  Classic Teams installation not found for user $($User.Name)"
    }
}

# Remove old Teams folders and icons
$TeamsFolder_old = "$($ENV:SystemDrive)\Users\*\AppData\Local\Microsoft\Teams"
$TeamsIcon_old = "$($ENV:SystemDrive)\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams*.lnk"
Get-Item $TeamsFolder_old | Remove-Item -Force -Recurse
Get-Item $TeamsIcon_old | Remove-Item -Force -Recurse

# Uninstall old New Teams Versions
Write-Host "Uninstalling previous version of New Microsoft Teams" -ForegroundColor Cyan
try {
    $Path = Join-Path -Path $DownloadDirectory -ChildPath $TeamsBootstrapperFileName
    Start-Process -FilePath  $Path -ArgumentList "-x" -NoNewWindow -Wait
    Write-Host "Uninstalled previous version of Microsoft Teams" -ForegroundColor Green
} catch {
    Write-Host "Error uninstalling previous version of Microsoft Teams" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Install WebRTC MSI
Write-Host "Installing AVD Remote Desktop WebRTC Redirector Service" -ForegroundColor Cyan
try {
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", "$DownloadDirectory\$WebRTCFileName", "Reboot=ReallySuppress", "/qn" -Wait -ErrorAction Stop
    Write-Host "Installed AVD Remote Desktop WebRTC Redirector Service" -ForegroundColor Green
} catch {
    Write-Host "Error installing AVD Remote Desktop WebRTC Redirector Service" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Überprüfung der installierten Version von Microsoft FSLogix Apps

# Definieren der erforderlichen Mindestversion
$requiredVersion = [version]"2.9.8784.63912"

# Funktion zur Vergleich der Versionen
function Compare-Versions {
    param (
        [version]$installedVersion,
        [version]$requiredVersion
    )

    if ($installedVersion -lt $requiredVersion) {
        Write-Warning "Die installierte Version ($installedVersion) von Microsoft FSLogix Apps ist niedriger als die erforderliche Version ($requiredVersion)."
    } else {
        Write-Host "Die installierte Version ($installedVersion) von Microsoft FSLogix Apps ist kompatibel." -ForegroundColor Green
    }
}

# Pfad zum Registry-Eintrag von FSLogix
$regPath = "HKLM:\SOFTWARE\FSLogix\Apps"

# Prüfen, ob der Registry-Eintrag vorhanden ist
if (Test-Path $regPath) {
    # Abrufen der installierten Version aus der Registry
    $installedVersion = (Get-ItemProperty -Path $regPath).InstallVersion

    if ($installedVersion) {
        $installedVersion = [version]$installedVersion
        Compare-Versions -installedVersion $installedVersion -requiredVersion $requiredVersion
    } else {
        Write-Warning "Die Version von Microsoft FSLogix Apps konnte nicht abgerufen werden."
    }
} else {
    Write-Warning "Microsoft FSLogix Apps ist nicht installiert."
}

# Install New Teams
Write-Host "Installing New Microsoft Teams" -ForegroundColor Cyan
try {
    $Path = Join-Path -Path $DownloadDirectory -ChildPath $TeamsBootstrapperFileName
    $TeamsInstallerPath = Join-Path -Path $DownloadDirectory -ChildPath $TeamsMSIXFileName
    Start-Process -FilePath $Path -ArgumentList "-p", "-o", "$TeamsInstallerPath" -Wait -ErrorAction Stop
    Write-Host "Installed new Microsoft Teams" -ForegroundColor Green
} catch {
    Write-Host "Error installing new Microsoft Teams" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Uninstall Teams Meeting Add-in
Write-Host "Uninstalling Teams Meeting Add-in" -ForegroundColor Cyan
try {
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/X{A7AB73A3-CB10-4AA5-9D38-6AEFFBDE4C91} /qn" -Wait -ErrorAction Stop
    Write-Host "Uninstalled Teams Meeting Add-in" -ForegroundColor Green
} catch {
    Write-Host "Error uninstalling Teams Meeting Add-in" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Disable new Teams auto update
Write-Host "Disabling new Teams auto update" -ForegroundColor Cyan
try {
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Teams" -Name disableAutoUpdate -PropertyType DWORD -Value 1 -Force
    Write-Host "Disabled new Teams auto update" -ForegroundColor Green
} catch {
    Write-Host "Error disabling new Teams auto update" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Install Teams Meeting Add-in
Write-Host "Installing Teams Meeting Add-in" -ForegroundColor Cyan
try {
    $TeamsAddinMSIPath = Get-ChildItem -Path 'C:\Program Files\WindowsApps' -Filter 'MSTeams*' -Directory | 
    ForEach-Object { 
        $_.FullName | Get-ChildItem -Filter '*.msi' -ErrorAction SilentlyContinue 
    } | Select-Object -ExpandProperty FullName
    $output = Get-AppLockerFileInformation -Path $TeamsAddinMSIPath | Select -ExpandProperty Publisher | select BinaryVersion
    $version = $output.BinaryVersion
    $major = $version.MajorPartNumber
    $minor = $version.MinorPartNumber
    $build = $version.BuildPartNumber
    $private = $version.PrivatePartNumber

    $teamsaddinversion = "$major.$minor.$build.$private"
    $teamsaddinversiopath = Join-Path -Path "C:\Program Files (x86)\Microsoft\TeamsMeetingAddin" -ChildPath $teamsaddinversion
    New-Item -Path $teamsaddinversiopath -ItemType Directory -Force > $null
    
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$TeamsAddinMSIPath`" ALLUSERS=1 /qn /norestart TARGETDIR=`"$teamsaddinversiopath`"" -Wait -ErrorAction Stop
    Write-Host "Installed Teams Meeting Add-in" -ForegroundColor Green
} catch {
    Write-Host "Error installing Teams Meeting Add-in" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Add Registry Keys for loading the Add-in
Write-Host "Adding Registry Keys for Teams Add-in" -ForegroundColor Cyan
try {
    New-Item -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins" -Name "TeamsAddin.FastConnect" -Force -ea SilentlyContinue
    New-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Type "DWord" -Name "LoadBehavior" -Value 3 -Force -ea SilentlyContinue
    New-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Type "String" -Name "Description" -Value "Microsoft Teams Meeting Add-in for Microsoft Office" -Force -ea SilentlyContinue
    New-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Type "String" -Name "FriendlyName" -Value "Microsoft Teams Meeting Add-in for Microsoft Office" -Force -ea SilentlyContinue

    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Type "DWord" -Name "EnableFullTrustStartupTasks" -Value "2" -Force -ea SilentlyContinue
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Type "DWord" -Name "EnableUwpStartupTasks" -Value "2" -Force -ea SilentlyContinue
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Type "DWord" -Name "SupportFullTrustStartupTasks" -Value "1" -Force -ea SilentlyContinue
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Type "DWord" -Name "SupportUwpStartupTasks" -Value "1" -Force -ea SilentlyContinue

    Write-Host "Added Registry Keys for Teams Add-in" -ForegroundColor Green
} catch {
    Write-Host "Error adding Registry Keys for Teams Add-in" -ForegroundColor Red
    Write-Host "Error: $($Error[0].Exception.Message)" -ForegroundColor Red
}

# Stop transcript logging
Stop-Transcript

exit