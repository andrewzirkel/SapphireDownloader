# SapphireDownloader
A Powershell Module to import reports from Sapphire SIS

## Installation:
Download zip file or use git: `git clone https://github.com/andrewzirkel/SapphireDownloader.git`

## Usage:
Use `Set-SDParameters` to set options

See examples in SapphireDownloaderExample.ps1

## Available Modules:
### Set-SDParameters

Set options for download.

`Set-SDParameters -Username $SDusername -Password $SDpassword -URL $sapphireURL -DistrictID $SDdistrict_id -SchoolID $SDschoold_id -SchoolYear $SDschool_year`

### Get-SDClass_Roster

Get Class Roster Report

### Get-SDDEMO_CUST_LIST

Get Custom Demographics report

### Get-SDReport (not working)

Get report writier report

### Get-SDConnectEd

Get ConnectEd (School Messenger) Report
