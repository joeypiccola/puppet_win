[CmdletBinding()]
Param(
$PSWindowsUpdateURL,
$WSUSscnURL
)

#$PSWindowsUpdateURL = 'http://nuget.ad.piccola.us:8081/pswindowsupdate.zip'
#$WSUSscnURL = 'http://nuget.ad.piccola.us:8081/wsusscn2.cab'


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
        # manually install the module based on what version of powershell is installed
        <#
        switch ($PSVersionTable.PSVersion.Major) {
            {($_ -eq '3') -or ($_ -eq '4')} {
                # update the PSModulePath with the new module directory if it does not exist
                if ($env:PSModulePath.Split(';') -notcontains 'C:\Program Files\WindowsPowerShell\Modules') {
                    New-Item -ItemType Directory -Path 'C:\Program Files\WindowsPowerShell\Modules' -Force -ErrorAction Stop
                    $CurrentValue = [Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
                    [Environment]::SetEnvironmentVariable("PSModulePath", $CurrentValue + ";C:\Program Files\WindowsPowerShell\Modules", "Machine")                    
                }
                # copy the contents of the unzipped module to the module directory. ignore versioning since PackageManagement and PowerShellGet or not being considered
                Copy-Item -Path 'C:\windows\Temp\pswindowsupdate\2.0.0.2' -Destination 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate' -Recurse -Force -ErrorAction Stop
            }
            '5' {
                # copy the contents of the unzipped module to the module directory.
                Copy-Item -Path 'C:\windows\Temp\pswindowsupdate\' -Destination 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate' -Recurse -Force -ErrorAction Stop            
            }
        }
        #>
    }
    # delete any existing wsusscn2.cab
    if (Test-Path -Path 'C:\Windows\temp\wsusscn2.cab') {
        Remove-Item -Path 'C:\Windows\temp\wsusscn2.cab' -Force -Confirm:$false -ErrorAction Stop
    }
    # download the wsusscn2.cab
    Start-BitsTransfer -Source $WSUSscnURL -Destination 'C:\windows\Temp' -ErrorAction Stop
    # import the pswindowesupdate module
    Import-Module -Name 'c:\windows\temp\PSWindowsUpdate\2.0.0.2\PSWindowsUpdate.psd1' -ErrorAction Stop
    # add the previously downloaded wsusscn2.cab file as an Offline Sync Service Manager
    Add-WUServiceManager -ScanFileLocation C:\Windows\Temp\wsusscn2.cab -Confirm:$false -ErrorAction Stop
    # get the service ID of the previoulsy added Offline Sync Service Manager
    $offlineSyncService = Get-WUServiceManager -ErrorAction Stop | ?{$_.name -eq 'Offline Sync Service'}
    # get the missing updates using the Offline Sync Service Manager
    $missingUpdates = Get-WindowsUpdate -ServiceID $offlineSyncService.ServiceID -ErrorAction Stop
    # get the previously added Offline Sync Service Manager and remove it (doing this manually since Remove-WUServiceManager stopped working as of PSWindowsUpdate 2.0.0.2)
    $offlineServiceManager = Get-WUServiceManager | ?{$_.name -eq 'Offline Sync Service'}
    $objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $objService = $objServiceManager.RemoveService($offlineSyncService.ServiceID)
    Write-Host "Total missing updates: $($missingUpdates.count)" -ForegroundColor Green
} catch {
    Write-Error $_.Exception.Message
}