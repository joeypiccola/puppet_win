[CmdletBinding()]
Param(
$PSWindowsUpdateURL,
$WSUSscnURL
)

function Expand-ZIPFile($File, $Destination) {
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($Destination).copyhere($item)
    }
}

try {
    # import the BitsTransfer module
    Import-Module -Name BitsTransfer -ErrorAction Stop
    # if the PSWindowsUpdate module is not installed, manually install it based on what version of powershell is available
    if (!(Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction Stop)) {
        # download and extract the pswindowsupdate module if it hasn't already been done
        if (!(Test-Path -Path C:\windows\Temp\pswindowsupdate -ErrorAction Stop)) {
            # download the pswindowsupdate module, automatically overwrites if already present
            Start-BitsTransfer -Source $PSWindowsUpdateURL -Destination 'C:\windows\Temp' -ErrorAction Stop
            # unzip the module
            Expand-ZIPFile -File C:\Windows\Temp\pswindowsupdate.zip -Destination C:\Windows\Temp
            # remove the module zip file
            Remove-Item -Path C:\Windows\Temp\pswindowsupdate.zip -Force -ErrorAction Stop
        }
    }
    # delete any existing wsusscn2.cab
    if (Test-Path -Path 'C:\Windows\temp\wsusscn2.cab') {
        Remove-Item -Path 'C:\Windows\temp\wsusscn2.cab' -Force -Confirm:$false -ErrorAction Stop
    }
    # download the wsusscn2.cab
    Start-BitsTransfer -Source $WSUSscnURL -Destination 'C:\windows\Temp' -ErrorAction Stop
    # import the pswindowesupdate module
    Import-Module -Name 'c:\windows\temp\PSWindowsUpdate\2.0.0.2\PSWindowsUpdate.psd1' -ErrorAction Stop
    # get any previous Offline Service Managers and remove them manually
    $offlineServiceManagers = Get-WUServiceManager | ?{$_.name -eq 'Offline Sync Service'}
    if ($offlineServiceManagers) {
        foreach ($offlineServiceManager in $offlineServiceManagers) {
            $objServiceManager = $null
            $objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
            $objService = $objServiceManager.RemoveService($offlineServiceManager.ServiceID)
        }
    }
    # add the previously downloaded wsusscn2.cab file as an Offline Sync Service Manager
    Add-WUServiceManager -ScanFileLocation C:\Windows\Temp\wsusscn2.cab -Confirm:$false -ErrorAction Stop
    # get the service ID of the previously added Offline Service Manager
    $offlineServiceManager = Get-WUServiceManager -ErrorAction Stop | ?{$_.name -eq 'Offline Sync Service'}
    # get the missing updates using the previously added Offline Service Manager
    $missingUpdates = Get-WindowsUpdate -ServiceID $offlineServiceManager.ServiceID -ErrorAction Stop
    # get the previously added Offline Sync Service Manager and remove it
    $offlineServiceManager = Get-WUServiceManager | ?{$_.name -eq 'Offline Sync Service'}
    $objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $objService = $objServiceManager.RemoveService($offlineServiceManager.ServiceID)
    Write-Host "Total missing updates: $($missingUpdates.count)" -ForegroundColor Green
} catch {
    Write-Error $_.Exception.Message
    exit 1
}