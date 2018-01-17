[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$PSWindowsUpdateURL
    ,
    [Parameter()]
    [Switch]$PSWindowsUpdateForceDownload
    ,
    [Parameter(Mandatory)]
    [String]$WSUSscnURL
    ,
    [Parameter()]
    [Switch]$WSUSscnForceDownload
)


#region helperFunctions

function Expand-ZIPFile($File, $Destination) {
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items()) {
        $shell.Namespace($Destination).copyhere($item)
    }
}

function Get-WebFileLastModified($url) {
    $webRequest = [System.Net.HttpWebRequest]::Create($url);
    $webRequest.Method = "HEAD";
    $webResponse = $webRequest.GetResponse()
    $remoteLastModified = ($webResponse.LastModified) -as [DateTime] 
    $webResponse.Close()
    Write-Output $remoteLastModified
}

#endregion

try {
    # import the BitsTransfer module
    Import-Module -Name BitsTransfer -ErrorAction Stop

    # download the pswindowsupdate module
    if (!(Test-Path -Path 'C:\windows\Temp\pswindowsupdate') -or ($PSWindowsUpdateForceDownload -eq $true)) {
        # remove pswindowsupdate dir if it exists
        if (Test-Path -Path 'C:\windows\Temp\pswindowsupdate') {
            Remove-Item -Path 'C:\windows\Temp\pswindowsupdate' -Force -ErrorAction Stop
        }
        # download the pswindowsupdate module, automatically overwrite the zip if already present
        Start-BitsTransfer -Source $PSWindowsUpdateURL -Destination 'C:\windows\Temp' -ErrorAction Stop
        # unzip the module
        Expand-ZIPFile -File 'C:\Windows\Temp\pswindowsupdate.zip' -Destination 'C:\Windows\Temp'
        # remove the module zip file
        Remove-Item -Path 'C:\Windows\Temp\pswindowsupdate.zip' -Force -ErrorAction Stop
    }

    # download the wsusscn2.cab file
    if (!(Test-Path -Path 'c:\windows\temp\wsusscn2.cab') -or (($WSUSscnForceDownload -eq $true) -and (Test-Path -Path 'c:\windows\temp\wsusscn2.cab'))) {
        # remove the wsusscn2.cab if it exists
        if (Test-Path -Path 'c:\windows\temp\wsusscn2.cab') {
            Remove-Item -Path 'C:\Windows\temp\wsusscn2.cab' -Force -Confirm:$false -ErrorAction Stop
        }
        # download the wsusscn2.cab
        Start-BitsTransfer -Source $WSUSscnURL -Destination 'C:\windows\Temp' -ErrorAction Stop
    } else {
        $localwsusscnFile = (Get-Item -Path -Path 'c:\windows\temp\wsusscn2.cab').LastWriteTime
        $remoteWSUSscnFile = Get-WebFileLastModified -url $WSUSscnURL
        # if the wsusscn2.cab file in the webrepo does not match the local version then redownload it
        if ($localwsusscnFile -ne $remoteWSUSscnFile) {
            Remove-Item -Path 'C:\Windows\temp\wsusscn2.cab' -Force -Confirm:$false -ErrorAction Stop
            Start-BitsTransfer -Source $WSUSscnURL -Destination 'C:\windows\Temp' -ErrorAction Stop
        }
    }

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

    # parse the results and stage an external fact
    $updates = $missingUpdates | select kb, title, size, msrcseverity, @{Name="LastDeploymentChangeTime";Expression={$_.lastdeploymentchangetime.tostring("MM-dd-yyyy")}}
    $kbarray = @()
    $updates | %{$kbarray += $_.kb} | ConvertTo-Csv
    $windowsupdatereporting_col = @()

    $update_meta = [pscustomobject]@{
        missing_update_count = $updates.Count
        missing_update = $updates
        missing_update_kbs = $kbarray
    }

    $scan_meta = [pscustomobject]@{
        last_run_time = (Get-Date -Format "MM-dd-yyyy")
        wsusscn2_file_time = (Get-Item -Path 'C:\windows\Temp\wsusscn2.cab').lastwritetime.ToString("MM-dd-yyyy")
    }

    $meta = [pscustomobject]@{
        scan_meta = $scan_meta
        update_meta = $update_meta
    }

    $fact_name = [pscustomobject]@{
        updatereporting_win = $meta
    }

    $windowsupdatereporting_col += $fact_name
    $windowsupdatereporting_col | ConvertTo-Json -Depth 4 | Out-File 'C:\ProgramData\PuppetLabs\facter\facts.d\updatereporting_win.json' -Force

} catch {
    Write-Error $_.Exception.Message
    exit 1
}