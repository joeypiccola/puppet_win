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

  String $pswindowsupdateurl = undef,
  String $wsusscnurl = undef,
  Boolean $pswindowsupdateforcedownload = false,
  Boolean $wsusscnforcedownload = false,

){

#  exec { 'run_powershell':
#    command   => "write-output '${value}'",
#    provider  => powershell,
#    logoutput => true,
#  }
#
#  file { 'puppet_win_psfile':
#    ensure => 'present',
#    source => 'puppet:///modules/puppet_win/Test-Param.ps1',
#    path   => 'c:/windows/temp/Test-Param.ps1',
#    before => Exec['puppet_win_psexec'],
#  }
#
#  exec { 'puppet_win_psexec':
#    command   => "& C:\\windows\\temp\\Test-Param.ps1 -value ${value} -valuetwo ${valuetwo}",
#    provider  => 'powershell',
#    logoutput => true,
#  }

  file { 'puppet_win_stage_file':
    ensure => 'present',
    source => 'puppet:///modules/puppet_win/Invoke-WindowsUpdateReport.ps1',
    path   => 'c:/windows/temp/Invoke-WindowsUpdateReport.ps1',
    before => Exec['puppet_win_run_file'],
  }

  exec { 'puppet_win_run_file':
    command   => "& C:\\windows\\temp\\Invoke-WindowsUpdateReport.ps1 -pswindowsupdateurl ${pswindowsupdateurl} -wsusscnurl ${wsusscnurl} -pswindowsupdateforcedownload:${pswindowsupdateforcedownload} -wsusscnforcedownload:${wsusscnforcedownload}",
    provider  => 'powershell',
    logoutput => true,
  }

}
