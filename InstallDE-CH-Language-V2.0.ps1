#Original script from https://www.inthecloud247.com/install-an-additional-language-pack-on-windows-11-during-autopilot-enrollment/
#This script need to run as System Context and as 64bit PowerShell

<#
.SYNOPSIS
  Script to install langauge pack and change MUI langauge

.DESCRIPTION
    Script to install langauge package and set default language

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file Invoke-ChangeDefaultLanguage.ps1 

.NOTES
    Credit: #Original script from https://www.inthecloud247.com/install-an-additional-language-pack-on-windows-11-during-autopilot-enrollment/
    Version:        1.0.0
    Author:         Sandy Zeng
    Creation Date:  09.06.2024
    Updated:    
    Version history:
        1.0.0 - (09.06.2024) Script released
#>

function Write-LogEntry {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,
        [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = $LogFileName
    )
    # Determine log file location
    $LogFilePath = Join-Path -Path $env:ProgramData -ChildPath $("Microsoft\IntuneManagementExtension\Logs\$FileName")
	
    # Construct time stamp for log entry
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
	
    # Construct date for log entry
    $Date = (Get-Date -Format "MM-dd-yyyy")
	
    # Construct context for log entry
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
	
    # Construct final log entry
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
	
    # Add value to log file
    try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
        if ($Severity -eq 1) {
            Write-Verbose -Message $Value
        }
        elseif ($Severity -eq 3) {
            Write-Warning -Message $Value
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "Unable to append log entry to $LogFileName file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

# The language we want as new default. Language tag can be found here: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/available-language-packs-for-windows?view=windows-11#language-packs
$LPlanguage = "de-CH"

#Region Initialisations
$LogFileName = "Invoke-ChangeDefaultLanguage-$LPlanguage.log"

# As In some countries the input locale might differ from the installed language pack language, we use a separate input local variable.
# A list of input locales can be found here: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-input-locales-for-windows-language-packs?view=windows-11#input-locales
$InputlocaleRegion = "de-CH"

# Geographical ID we want to set. GeoID can be found here: https://learn.microsoft.com/en-us/windows/win32/intl/table-of-geographical-locations
$geoId = "223"

#Install language pack and change the language of the OS on different places
#Install an additional language pack including FODs. With CopyToSettings (optional), this will change language for non-Unicode program. 
Write-LogEntry -Value "Installing language $LPlanguage" -Severity 1
try {
    Install-Language -Language $LPlanguage -CopyToSettings -ErrorAction Stop
    Write-LogEntry -Value "$LPlanguage is installed" -Severity 1
}
catch [System.Exception] {
    Write-LogEntry -Value "$LPlanguage install failed with error: $($_.Exception.Message)" -Severity 3
    exit 1
}

# Configure new language defaults under current user (system) after which it can be copied to system
Write-LogEntry -Value "Set Win UI Language Override for regional changes $InputlocaleRegion " -Severity 1
Set-WinUILanguageOverride -Language $InputlocaleRegion

# adding the input locale language to the preferred language list, and make it as the first of the list. 
Write-LogEntry -Value "Set Win User Language List" -Severity 1
$OldList = Get-WinUserLanguageList
$UserLanguageList = New-WinUserLanguageList -Language $InputlocaleRegion
$UserLanguageList += $OldList
Set-WinUserLanguageList -LanguageList $UserLanguageList -Force

# Set Win Home Location, sets the home location setting for the current user. This is for Region location 
Write-LogEntry -Value "Set Region location $geoId" -Severity 1
Set-WinHomeLocation -GeoId $geoId

# Set Culture, sets the user culture for the current user account. This is for Region format
Write-LogEntry -Value "Set Region format $InputlocaleRegion" -Severity 1
Set-Culture -CultureInfo $InputlocaleRegion

# Copy User International Settings from current user to System, including Welcome screen and new user
Write-LogEntry -Value "Copy User International Settings from current user to System" -Severity 1
Copy-UserInternationalSettingsToSystem -WelcomeScreen $True -NewUser $True

Exit 3010