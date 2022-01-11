# CHANGELOG for Microsoft.Powershell_Profile.ps1

## Jan 10, 2022

- Fixed jit policy; changed back to 5 hours or more
- Added VM Manager function; loops vms with limited resources
- Added all switch to module updater; can be ran to clean multiple copies of a module
## Dec 20, 2021

- Fixed multiple Azure login; kept trying to use previous login
- Fixed NSG retrieval; conflicted with Bastion Host NSG's
- Updated Module display; to many spaces caused wrapping text.
- added NoRun switch to allow quicker runtime

## Oct 05, 2021

- Added -SkipPublisherCheck to module updater; if new modules were updated from base; it allows install like Pester
- Corrected misspelled outputs and resolved aliases; fixed most lint issues

## Sep 29, 2021

- Fixed Start-MyAzureVM pre-populated VMname; was coming back with object and not name

## Sep 6, 2021

- Added resources group parameter to Start-AZVM commands
- Added auto fill for VMname when using parameter and tab
- Removed unused commands; not needed for Azure lab
- removed leading spaces.

## Aug 28, 2021

- changed functions to standard my format
- Fixed Azure connection prompting for Resource group all the time; added global variable check
- Added Start-MyLabEnvironment function; set a change of commands for other functions automates entire start
- Added ConnectTO-MyAzureVM; added ability to select public IP or DNS. Credential param still do not work

## Jun 11, 2021

- changed functions to standard my format
- fixed Voice to female for both Windows 10 and 11; added more comment by voice with voice parameter for some commands
- Renamed all functions to incorporate 'my' in function; resolved conflict with other functions

## May 19, 2021

- updated code to be more responsive
- Changed Tenant variable to global variables; Fixed multiple functions from using same variables
- Moved important variable to top of script
- Fixed Function Get-AzureUsername to look at office identity instead of Store identity

## Apr 9, 2021

- Removed unneeded Modules; just specified the main ones
- Added VScode installer and check
- Added support functions for commands

## Feb 22, 2021

- Renamed Function Prep-MyAzureEnvironment to Start-MyAzureEnvironment
- Changed JIT policy to 3 hrs; matches Azure policy

## Nov 17, 2020

- Initial
