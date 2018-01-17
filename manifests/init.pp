# == Class: puppet_win
#
#
#    
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
  Boolean $pswindowsupdateforcedownload = false,
  Boolean $wsusscnforcedownload = false,

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
    before => Exec['puppet_win_run_file'],
  }

  exec { 'puppet_win_run_file':
    command   => "& C:\\windows\\temp\\Invoke-WindowsUpdateReport.ps1 -pswindowsupdateurl ${pswindowsupdateurl} -wsusscnurl ${wsusscnurl} -pswindowsupdateforcedownload:${pswindowsupdateforcedownload_set} -wsusscnforcedownload:${wsusscnforcedownload_set}",
    provider  => 'powershell',
    logoutput => true,
  }

}
