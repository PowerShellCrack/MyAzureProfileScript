# CHANGELOG for Microsoft.Powershell_Profile.ps1


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