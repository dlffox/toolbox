<#

    .SYNOPSIS
    Collect network info.

    .DESCRIPTION
    Collects usefull info about active network adapters.

    .NOTES
    Version:        1.0
    Author:         John Fox Maule
    Creation date:  2025/08/27

#>

$ExternalNameServers = @('1.1.1.1', '8.8.8.8', '9.9.9.9')
$MediaTypes = @('802.3', 'Native 802.11')

$DesktopPath = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'Desktop'
$ChildPath = 'Network-{0}-{1}.log' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMddTHHmmssZ')
$JoinedPath = Join-Path -Path $DesktopPath.Desktop -ChildPath $ChildPath
$TempFile = New-TemporaryFile

Start-Transcript -Path $JoinedPath

Write-Output '#--------------------------------------------------------------------------------'
Write-Output '#  Timezone'
Write-Output '#--------------------------------------------------------------------------------'

Get-TimeZone

# Get active adapters
Write-Output '#--------------------------------------------------------------------------------'
Write-Output '#  Getting active network adapters'
Write-Output '#--------------------------------------------------------------------------------'

Get-NetAdapter | Where-Object { $_.MediaType -In $MediaTypes -and $_.Status -eq 'Up' } | Select-Object Name, InterfaceDescription, ifIndex, MacAddress, LinkSpeed | Format-Table
$NetAdapters = Get-NetAdapter | Where-Object { $_.MediaType -In $MediaTypes -and $_.Status -eq 'Up' }

$NetAdapters | ForEach-Object { 
    Write-Output '#--------------------------------------------------------------------------------'
    Write-Output ('#  Getting adapter properties for ''{0}''' -f $_.Name)
    Write-Output '#--------------------------------------------------------------------------------'
    
    Get-NetAdapterAdvancedProperty -Name $_.Name | Select-Object DisplayName, DisplayValue, RegistryKeyword, RegistryValue | Format-Table
}

$NetAdapters | ForEach-Object {
    Write-Output '#--------------------------------------------------------------------------------'
    Write-Output ('#  Getting IP configuration for ''{0}''' -f $_.Name)
    Write-Output '#--------------------------------------------------------------------------------'

    Get-NetAdapter -Name $_.Name | Get-NetIPConfiguration -Detailed
}


# Get DNS info and try resolving name using supplied and external dns-server
Write-Output '#--------------------------------------------------------------------------------'
Write-Output '#  Getting global dns settings'
Write-Output '#--------------------------------------------------------------------------------'

Get-DnsClientGlobalSetting

Write-Output '#--------------------------------------------------------------------------------'
Write-Output '#  Resolving dns name using supplied dns server'
Write-Output '#--------------------------------------------------------------------------------'

Resolve-DnsName -Name www.google.com -Type A

Write-Output '#--------------------------------------------------------------------------------'
Write-Output '#  Resolving dns name using external dns servers'
Write-Output '#--------------------------------------------------------------------------------'

foreach ($ExternalNameServer in $ExternalNameServers) {
    Resolve-DnsName -Name dkaag01-fw01.bugfinder.dk -Type A -Server $ExternalNameServer
}

# There is no equivalent powershell cmdlet at least not to my knowledge
Invoke-Command -ScriptBlock { netsh wlan show all > $TempFile }
Get-Content -Path $TempFile

Remove-Item -Path $TempFile

Stop-Transcript