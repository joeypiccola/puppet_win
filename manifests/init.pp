# == Class: puppet_win
#
#
#    
# exec { 'puppet_win_run_file':
#   command   => "& C:\\windows\\temp\\Invoke-WindowsUpdateReport.ps1 -pswindowsupdateurl ${pswindowsupdateurl} -wsusscnurl ${wsusscnurl} -pswindowsupdateforcedownload:${pswindowsupdateforcedownload_set} -wsusscnforcedownload:${wsusscnforcedownload_set} -downloaddirectory ${downloaddirectory}",
#   provider  => 'powershell',
#   logoutput => true,
# }
#
# === Parameters
#
# [*value*]
#   The timezone to use. For a full list of available timezone run tzutil /l.
#   Use the listed time zone ID (e.g. 'Eastern Standard Time')
#
# === Examples
#
#  class { ::puppet_win:
#    timezone = 'Mountain Standard Time',
#  }
#
# === Authors
#
# Joey Piccola <joey@joeypiccola.com>
#
# === Copyright
#
# Copyright (C) 2016 Joey Piccola.
#
class puppet_win (

  String $pswindowsupdateurl,
  String $wsusscnurl,
  String $downloaddirectory = 'c:/Windows/Temp',
  Boolean $pswindowsupdateforcedownload = false,
  Boolean $wsusscnforcedownload = false,
  String $dayofweek = 'sun',

){

  case $pswindowsupdateforcedownload {
    true: {
      $pswindowsupdateforcedownload_set = '$true'
  }
    default: {
      $pswindowsupdateforcedownload_set = '$false'
    }
  }

  case $wsusscnforcedownload {
    true: {
      $wsusscnforcedownload_set = '$true'
  }
    default: {
      $wsusscnforcedownload_set = '$false'
    }
  }

  file { 'puppet_win_stage_file':
    ensure => 'present',
    source => 'puppet:///modules/puppet_win/Invoke-WindowsUpdateReport.ps1',
    path   => 'c:/windows/temp/Invoke-WindowsUpdateReport.ps1',
    before => Scheduled_task['updatereporting_win'],
  }

  $min = fqdn_rand(59)
  $hour = fqdn_rand(3)+1

  scheduled_task { 'updatereporting_win':
    ensure    => 'present',
    name      => 'Windows Update Reporting (Puppet Managed Scheduled Task)',
    enabled   => true,
    command   => 'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe',
    arguments => "-WindowStyle Hidden -ExecutionPolicy Bypass \"C:\\windows\\temp\\Invoke-WindowsUpdateReport.ps1 -pswindowsupdateurl ${pswindowsupdateurl} -wsusscnurl ${wsusscnurl} -pswindowsupdateforcedownload:${pswindowsupdateforcedownload_set} -wsusscnforcedownload:${wsusscnforcedownload_set} -downloaddirectory ${downloaddirectory}\"",
    provider  => 'taskscheduler_api2',
    trigger   => {
      schedule    => weekly,
      day_of_week => $dayofweek,
      start_time  => "${hour}:${min}",
    }
  }
}
