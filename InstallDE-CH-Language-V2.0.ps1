# ------------------------------------------------------------------------------------------------------------ #
# Author(s)    : Peter Klapwijk - www.inthecloud247.com                                                        #
#                (Part of the script is from script from www.oliverkieselbach.com)                             #
# Version      : 1.1 (hardened for unattended Intune deployment)                                               #
#                                                                                                              #
# Description  : Install an additional language pack including FODs.                                           #
#                Changes language for new users, welcome screen etc.                                           #
#                Runs fully unattended (SYSTEM context), also when no user is logged on.                       #
#                Supported on Windows 11 22H2 and later.                                                       #
# ------------------------------------------------------------------------------------------------------------ #

# Microsoft Intune Management Extension might start a 32-bit PowerShell instance. If so, restart as 64-bit PowerShell
If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
    }
    Catch {
        Throw "Failed to start $PSCOMMANDPATH"
    }
    Exit
}

# Set variables:
# Company name
$CompanyName = "VISI"
# The language we want as new default.
$language = "de-CH"
# Geographical ID we want to set.
$geoId = "223"  # Swiss

# Start Transcript
Start-Transcript -Path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$($(Split-Path $PSCommandPath -Leaf).ToLower().Replace(".ps1",".log"))" | Out-Null

# custom folder for temp scripts
"...creating custom temp script folder"
$scriptFolderPath = "$env:SystemDrive\ProgramData\$CompanyName\CustomTempScripts"
New-Item -ItemType Directory -Force -Path $scriptFolderPath | Out-Null
"`n"

$userConfigScriptPath = $(Join-Path -Path $scriptFolderPath -ChildPath "UserConfig.ps1")
"...creating userconfig script"
$userConfigScript = @"
`$language = "$language"

Start-Transcript -Path "`$env:TEMP\LXP-UserSession-Config-`$language.log" | Out-Null

`$geoId = $geoId

# important for regional change like date and time...
"Set-WinUILanguageOverride = `$language"
Set-WinUILanguageOverride -Language `$language

"Set-WinUserLanguageList = `$language"

`$OldList = Get-WinUserLanguageList
`$UserLanguageList = New-WinUserLanguageList -Language `$language
`$UserLanguageList += `$OldList | where { `$_.LanguageTag -ne `$language }
"Setting new user language list:"
`$UserLanguageList | select LanguageTag
""
"Set-WinUserLanguageList -LanguageList ..."
Set-WinUserLanguageList -LanguageList `$UserLanguageList -Force

"Set-Culture = `$language"
Set-Culture -CultureInfo `$language

"Set-WinHomeLocation = `$geoId"
Set-WinHomeLocation -GeoId `$geoId

Stop-Transcript -Verbose
"@

# Install an additional language pack including FODs (with one retry for transient download issues)
"Installing language pack $language"
$languageInstalled = $false
$attempt = 0
while (-not $languageInstalled -and $attempt -lt 2) {
    $attempt++
    try {
        Install-Language $language -CopyToSettings -ErrorAction Stop
        $languageInstalled = $true
    }
    catch {
        "Attempt $attempt to install language failed: $($_.Exception.Message)"
        if ($attempt -lt 2) { Start-Sleep -Seconds 20 }
    }
}

# Check status of the installed language pack
"Checking installed language pack status"
$installedLanguage = (Get-InstalledLanguage).LanguageId
if ($installedLanguage -like $language) {
    Write-Host "Language $language installed"
}
else {
    Write-Host "Failure! Language $language NOT installed"
    Stop-Transcript
    Exit 1
}

# Set System Preferred UI Language
"Set SystemPreferredUILanguage"
Set-SystemPreferredUILanguage $language

# Check status of the System Preferred Language
$SystemPreferredUILanguage = Get-SystemPreferredUILanguage
if ($SystemPreferredUILanguage -like $language) {
    Write-Host "System Preferred UI Language set to $language. OK"
}
else {
    Write-Host "Failure! System Preferred UI Language NOT set to $language. System Preferred UI Language is $SystemPreferredUILanguage"
    Stop-Transcript
    Exit 1
}

# Configure new language defaults under current user (SYSTEM account) so they can be copied to the system
"Set WinUILanguageOverride"
Set-WinUILanguageOverride -Language $language

"Set WinUserLanguageList"
$OldList = Get-WinUserLanguageList
$UserLanguageList = New-WinUserLanguageList -Language $language
$UserLanguageList += $OldList | where { $_.LanguageTag -ne $language }
$UserLanguageList | select LanguageTag
Set-WinUserLanguageList -LanguageList $UserLanguageList -Force

"Set culture"
Set-Culture -CultureInfo $language

"Set WinHomeLocation"
Set-WinHomeLocation -GeoId $geoId

# Copy User International Settings from current user (SYSTEM) to System, incl. Welcome screen and new user
"Copy UserInternationalSettingsToSystem"
try {
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $True -NewUser $True -ErrorAction Stop
}
catch {
    "Copy-UserInternationalSettingsToSystem failed (non-fatal): $($_.Exception.Message)"
}

# Switch the language for the CURRENT user session, but only if a user is logged on.
# When running in device context (e.g. Autopilot ESP without user), this block is skipped.
$loggedOnUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
if ([string]::IsNullOrWhiteSpace($loggedOnUser)) {
    "No interactive user logged on. Skipping current-user session config."
    "New user, Welcome screen and system defaults are already configured via Copy-UserInternationalSettingsToSystem."
}
else {
    "Interactive user detected: $loggedOnUser. Triggering language change for current user session."
    Out-File -FilePath $userConfigScriptPath -InputObject $userConfigScript -Encoding ascii

    $taskName = "LXP-UserSession-Config-$language"
    # Run PowerShell directly (no wscript/VBS) to avoid AppLocker/ASR blocking and window flashing
    $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$userConfigScriptPath`""
    $trigger   = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -UserId $loggedOnUser -LogonType Interactive
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    $task      = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
    Register-ScheduledTask $taskName -InputObject $task -Force | Out-Null
    Start-ScheduledTask -TaskName $taskName

    # Wait for the task to finish instead of a fixed sleep
    $maxWaitSeconds = 180
    $waited = 0
    Start-Sleep -Seconds 3
    while (((Get-ScheduledTask -TaskName $taskName).State -eq 'Running') -and ($waited -lt $maxWaitSeconds)) {
        Start-Sleep -Seconds 5
        $waited += 5
    }
    $lastResult = (Get-ScheduledTaskInfo -TaskName $taskName).LastTaskResult
    "Current-user task finished with LastTaskResult $lastResult after about $waited seconds"

    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# trigger 'LanguageComponentsInstaller\ReconcileLanguageResources' otherwise 'Windows Settings' need a long time to change finally
"Trigger ScheduledTask = LanguageComponentsInstaller\ReconcileLanguageResources"
Start-ScheduledTask -TaskName "\Microsoft\Windows\LanguageComponentsInstaller\ReconcileLanguageResources"

Start-Sleep 10

# trigger store updates, there might be new app versions due to the language change
"Trigger MS Store updates for app updates"
Get-CimInstance -Namespace "root\cimv2\mdm\dmmap" -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" | Invoke-CimMethod -MethodName "UpdateScanMethod"

# Add registry key for Intune detection
"Add registry key for Intune detection"
REG add "HKLM\Software\$CompanyName\LanguageXPWIN11\v1.0" /v "SetLanguage-$language" /t REG_DWORD /d 1 /f

Stop-Transcript
Exit 0