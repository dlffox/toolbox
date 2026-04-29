<#

    .SYNOPSIS
    Perform an packet capture.

    .DESCRIPTION
    Perform an packet capture, and convert to pcap next generation format for use with Wireshark.

    .PARAMETER OutputPath
    Specifies the directory path where the the PcapNextGeneration file will be stored.

    .INPUTS
    None. You can't pipe objects to PacketCapture.

    .EXAMPLE
    PacketCapture -OutputPath C:\temp

    .NOTES
    Author:         John Fox Maule
    Creation date:  20260429
    Version:        1.0

    .LINK
    https://learn.microsoft.com/en-us/windows-server/networking/technologies/pktmon/pktmon

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]
    $OutputPath = 'C:\temp'
)

#--------------------------------------------------------------------------------
#
#  Functions 
#
#--------------------------------------------------------------------------------

function ConvertTo-PcapNextGeneration  {
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter()]
        [string]
        $InputFile,

      # Parameter help description
        [Parameter(Mandatory=$true)]
        [string]
        $PcapNextGenerationFile
    )

    if (Test-Path -Path $InputFile) {
        $ArgumentList = ('etl2pcap {0} --out {1}' -f $InputFile, $PcapNextGenerationFile)
        Start-Process -FilePath 'PktMon.exe' -ArgumentList $ArgumentList -Wait
    }

    Write-Host "🦊 says: " -NoNewline -ForegroundColor DarkGreen && Write-Host ('Collect your wire🦈 file here {0}' -f $PcapNextGenerationFile) -ForegroundColor White
}
function Initialize-PacketCapture {

    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory)]
        [string]
        $OutputPath
    )

    try {

        # ¿Are we running as admin?
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            throw '🦊 says: Please run from elevated PowerShell prompt'
        }
        
        # Lets make sure we have the output directory.
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }

    }
    catch {
        Write-Error -Message $PSItem.Exception.Message
        throw
    }
}

function Invoke-PacketCapture {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $PcapNextGenerationFile
    )

    # Lets create some workfiles.
    $CaptureFile = New-CaptureFile
    $TemporaryFile = New-TemporaryFile

    try {

        # Check if PktMon is already running.
        Start-Process -FilePath 'PktMon.exe' -ArgumentList 'status' -RedirectStandardOutput $TemporaryFile
        $PktMonNotRunning = Select-String -Pattern 'Packet Monitor is not running.' -Path $TemporaryFile
        if ($null -ne $PktMonNotRunning) {
            Stop-PacketCapture
        }
        
        # I want to see everything, hence remove all filter and set pkt-size to zero.
        Start-Process -FilePath 'PktMon.exe' -ArgumentList 'filter remove' -RedirectStandardOutput $TemporaryFile
        $ArgumentList = ('start --capture --pkt-size 0 --file-name {0}' -f $CaptureFile) 
        Start-Process -FilePath 'PktMon.exe' -ArgumentList $ArgumentList -RedirectStandardOutput $TemporaryFile -NoNewWindow

        # ¿Self-explanatory? 
        Read-Host -Prompt 'Press enter when you have reproduced the problem.'
        Stop-PacketCapture
        ConvertTo-PcapNextGeneration -InputFile $CaptureFile -PcapNextGenerationFile $PcapNextGenerationFile
    }

    catch {
        Write-Error -Message $PSItem.ErrorDetails.Message
    }
    finally {
        if (Test-Path -Path $CaptureFile) {
            Remove-Item -Path $CaptureFile
        }
        if (Test-Path -Path $TemporaryFile) {
            Remove-Item -Path $TemporaryFile
        }
    }    
}

function New-CaptureFile {
    
    $CaptureFile = New-TemporaryFile
    $NewName = ('{0}.etl' -f $CaptureFile.BaseName)
    Rename-Item -Path $CaptureFile.FullName -NewName $NewName

    return Join-Path -Path $CaptureFile.DirectoryName -ChildPath $NewName
}

function Stop-PacketCapture {
    Start-Process -FilePath 'PktMon.exe' -ArgumentList 'stop'
}

#--------------------------------------------------------------------------------
#
#  Main 
#
#--------------------------------------------------------------------------------

$BaseName = Read-Host -Prompt 'Please specify capture name eg. ticket id'
Initialize-PacketCapture -OutputPath $OutputPath
$PcapNextGenerationFile = Join-Path -Path $OutputPath -ChildPath ('{0}.pcapng' -f $BaseName)
Invoke-PacketCapture -PcapNextGenerationFile $PcapNextGenerationFile