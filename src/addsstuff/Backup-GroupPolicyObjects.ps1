<#

    .SYNOPSIS
    Backup all the GPOs in a domain.

    .DESCRIPTION
    Backup all the GPOs in a domain to a specified path and creates a html report.

    .PARAMETER Cleanup
    If the switch is set, it will cleanup old backups that are older than 90 days.

    .PARAMETER Domain
    Specifies the fully qualified domain name (FQDN) of the active directory services domain.

    .PARAMETER Path
    Specifies the path to the backup directory.

    .EXAMPLE
    Backup-GroupPolicyObjects -Domain 'bugfinder.dk' -Path 'C:\GPOBackups'

    .EXAMPLE
    Backup-GroupPolicyObjects -Domain 'bugfinder.dk' -Path 'C:\GPOBackups' -Cleanup

    .INPUTS
    None. You can't pipe objects to Backup-GroupPolicyObjects.

    .OUTPUTS
    None.

    .NOTES
    Author:         John Fox Maule 🦊
    Creation date:  20260511
    Version:        1.0

    .LINK
    https://learn.microsoft.com/en-us/powershell/module/grouppolicy/?view=windowsserver2022-ps

    .LINK
    https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wmi_v2/class-library/gpo-class-microsoft-grouppolicy

    .LINK
    Backup-GPO

    .LINK
    Get-GPO

#>

#-------------------------------------------------------------------------------
#
#  Parameters 
#
#-------------------------------------------------------------------------------

[CmdletBinding()]
param (    
    # If the switch is set, it will cleanup old backups.
    [Parameter(Mandatory=$false)]
    [switch]
    $Cleanup = $false,

    # Specifies the fully qualified domain name (FQDN) of the active directory services domain.
    [Parameter(Mandatory=$true)]
    [string]
    $Domain,

    # Specifies the path to the backup directory.
    [Parameter(Mandatory=$true)]
    [string]
    $Path
)

#-------------------------------------------------------------------------------
#
#  Functions 
#
#-------------------------------------------------------------------------------

function Add-LogEntry {

    [CmdletBinding()]
    param (
        # The level of the log entry. It can be Debug, Error, Information or Warning. The default is Information.
        [Parameter(Mandatory=$false)]
        [string]
        [ValidateSet('Debug', 'Error', 'Information', 'Warning')]
        $Level = 'Information',

        # The message of the log entry. It is mandatory.
        [Parameter(Mandatory)]
        [string]
        $Message
    )

    # Initialize variables.
    $LogFile = Get-Variable -Name LogFile -Scope Script -ValueOnly

    # Log the message to the console based on the level and also log it to the context if the level is not Debug.
    switch ($Level) {
        'Debug' { Write-Debug -Message $Message }
        'Error' { Write-Error -Message $Message }
        'Information' { Write-Information -MessageData $Message }
        'Warning' { Write-Warning -Message $Message }
        default { Write-Information -MessageData $Message }
    }

    # Log the message to the context if the level is Error.
    if ($Level -eq 'Error') {
        # $Context.LogException($Message)
    }

    # Log the message to the context if the level is Information or Warning.
    if ($Level -eq 'Information' -or $Level -eq 'Warning') {
        # $Context.LogMessage($Message, $Level)
    }

    try {
        # Log the message to the log file. If it fails, it will throw an error and we will catch it and continue with the rest of the script.
        Add-Content -Path $LogFile -Value ('[{0}Z] {1}: {2}' -f (Get-Date -Format 'yyyyMMddTHHmmss'), $Level, $Message) -ErrorAction Stop
    }
    catch {
        throw 'Failed to log message to the log file. Please check the path and your permissions.'
    }
}

function Backup-GroupPolicyObjects {

    # Initialize variables.
    $BackupPath = Get-Variable -Name BackupPath -Scope Script -ValueOnly
    $Domain = Get-Variable -Name Domain -Scope Script -ValueOnly
    $GPOHashTable = @{}
    $NotBackedUp = 0
    $SuccessfullyBackedUp = 0
    
    try {
        # Get all GPOs in the domain.
        $GroupPolicyObjects = Get-GPO -All -Domain $Domain -ErrorAction Stop
    }
    catch {
        Add-LogEntry -Message 'Failed to get GPOs from the domain. Please check the domain name and your permissions.' -Level Error
        throw 'Failed to get GPOs from the domain. Please check the domain name and your permissions.'
    }

    # Loop through each GPO and backup it up to the specified path.
    foreach ($GroupPolicyObject in $GroupPolicyObjects) {

        $CurrentGroupPolicyObject = [PSCustomObject]@{
            CreationTime = $GroupPolicyObject.CreationTime
            Description = $GroupPolicyObject.Description
            DisplayName = $GroupPolicyObject.DisplayName
            GpoStatus = $GroupPolicyObject.GpoStatus
            Id = $GroupPolicyObject.Id
            ModificationTime = $GroupPolicyObject.ModificationTime
            Owner = $GroupPolicyObject.Owner
            Path = $GroupPolicyObject.Path
        }

        # Add it to the hashtable for later use in the report.
        $GPOHashTable.Add($GroupPolicyObject.Id, $CurrentGroupPolicyObject)

        # Log the name of the GPO we are currently backing up.
        Add-LogEntry -Message ('------------------------------------------------------------------------------------------------') -Level Information

        try {
            # Backup the GPO to the specified path. If it fails, it will throw an error and we will catch it and continue with the next GPO.    
            Backup-GPO -Guid $GroupPolicyObject.Id -Path $BackupPath -Domain $Domain -ErrorAction SilentlyContinue | Out-Null
            Add-LogEntry -Message ('Backing up GPO: {0}...Succeeded' -f $GroupPolicyObject.DisplayName) -Level Information
            $SuccessfullyBackedUp++
        }
        catch {
            # If the backup fails, we will log it and continue with the next GPO.
            $NotBackedUp++
            Add-LogEntry -Message ('Backing up GPO: {0}...Failed' -f $GroupPolicyObject.DisplayName) -Level Error
        }
    }
    
    # Log the results of the backup process.
    Add-LogEntry -Message '------------------------------------------------------------------------------------------------' -Level Information
    Add-LogEntry -Message ('{0} GPOs were successfully backed up.' -f $SuccessfullyBackedUp) -Level Information
    if ($NotBackedUp -gt 0) {
        Add-LogEntry -Message ('{0} GPOs were not backed up.' -f $NotBackedUp) -Level Warning
    }   
    
    # Lets calculate some filepaths for the reports.
    $UtcNow = (Get-Date).ToUniversalTime()
    $Basefile = ('{0}Z - {1}' -f (Get-Date -Date $UtcNow -Format 'yyyyMMddTHHmmss'), $Domain)
    $ParentPath = Split-Path -Path $BackupPath -Parent -Resolve
    $CommaSeparatedPath = Join-Path -Path $ParentPath -ChildPath ('{0}.csv' -f $Basefile)
    $HTMLPath = Join-Path -Path $ParentPath -ChildPath ('{0}.htm' -f $Basefile)
    
    # Create a html report and a comma separated file for later use.
    $Title = ('Group policies backup report for {0} {1}Z' -f $Domain, (Get-Date -Format 'yyyyMMddTHHmmss'))
    $GPOHashTable.Values | ConvertTo-Html -Title $Title | Out-File -FilePath $HTMLPath
    $GPOHashTable.Values | Export-Csv -Path $CommaSeparatedPath -NoTypeInformation -Encoding utf8
}

function Initialize-GroupPolicyObjects {

    [CmdletBinding()]
    param (
        # Specifies the fully qualified domain name (FQDN) of the active directory services domain.
        [Parameter(Mandatory=$true)]
        [string]
        $Domain,

        # Specifies the path where the backups will be stored.
        [Parameter(Mandatory)]
        [string]
        $Path
    )

    # Lets calculate some filepaths .
    $Now = (Get-Date).ToUniversalTime()
    $Month = ('{0:d2}' -f $Now.Month)
    $Day = ('{0:d2}' -f $Now.Day)
    $YearPath = Join-Path -Path $Path -ChildPath $Now.Year
    $MonthPath = Join-Path -Path $YearPath -ChildPath $Month
    $DayPath = Join-Path -Path $MonthPath -ChildPath $Day
    $LogFile = Join-Path -Path $MonthPath -ChildPath ('{0} - {1}.log' -f (Get-Date -Date $Now -Format 'yyyyMMdd'), $Domain)
    
    # Initialize variables.
    Set-Variable -Name BackupPath -Scope Script -Value $DayPath -Visibility Private
    Set-Variable -Name BasePath -Scope Script -Value $Path -Visibility Private
    Set-Variable -Name Domain -Scope Script -Value $Domain -Visibility Private
    Set-variable -Name LogFile -Scope Script -Value $LogFile -Visibility Private

    # Check if the GroupPolicy module is available. If it is, we will import it. If it is not, we will log it and stop the script.
    if (Get-Module -Name GroupPolicy -ListAvailable) {
        Import-Module -Name GroupPolicy -Force
    }
    else {
        Add-LogEntry -Message 'GroupPolicy module is not available. Please install the module and try again.' -Level Error
        throw 'GroupPolicy module is not available. Please install the module and try again.'
    }

    # Create the directory structure for the backups. If it fails, we will log it and stop the script.
    try {
        # Create Path if it does not exist.
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }

        # Create YearPath if it does not exist.
        if (-not (Test-Path -Path $YearPath)) {
            New-Item -Path $YearPath -ItemType Directory -Force | Out-Null
        }
        
        # Create MonthPath if it does not exist.
        if (-not (Test-Path -Path $MonthPath)) {
            New-Item -Path $MonthPath -ItemType Directory -Force | Out-Null
        }

        # Create DayPath if it does not exist.
        if (-not (Test-Path -Path $DayPath)) {
            New-Item -Path $DayPath -ItemType Directory -Force | Out-Null
        }
    }
    catch {
        # If the initialization process fails, we will log it and stop the script.
        Add-LogEntry -Message $PSItem.ErrorDetails.Message -Level Error
        throw 'Failed to initialize the backup process. Please check the path and your permissions.'
    }
}

function Invoke-Cleanup {

    # Initialize variables.
    $CleanupPath = Get-Variable -Name BasePath -Scope Script -ValueOnly

    try {
        $Target = (Get-Date).ToUniversalTime().AddDays(-90)

        # Find files older than target date and remove them.
        Get-ChildItem -Path $CleanupPath -File -Recurse | Where-Object { $_.CreationTimeUtc -lt $Target } | Remove-Item

        # Find empty directories and remove them.
        Get-ChildItem -Path $CleanupPath -Directory -Recurse | Where-Object { $_.GetFileSystemInfos().Count -eq 0 } | Remove-item  
    }
    catch {
        # If the cleanup process fails, we will log it and continue with the rest of the script.
        Add-LogEntry -Message 'Failed to cleanup old backups. Please check the path and your permissions.' -Level Error
        Add-LogEntry -Message $PSItem.ErrorDetails.Message -Level Error
    }
}

#-------------------------------------------------------------------------------
#
#  Main 
#
#-------------------------------------------------------------------------------

Initialize-GroupPolicyObjects -Path $Path -Domain $Domain
Backup-GroupPolicyObjects

# If the cleanup switch is set, we will cleanup old backups that are older than 90 days.
if ($Cleanup.IsPresent) {
    Invoke-Cleanup
}