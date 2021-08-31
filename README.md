# MyAzureProfileScript

A Profile script for your admin system to manage VMs in multiple (or singular) Azure tenants. I have multiple tenants and need to start them up based on what I am testing.

## Customize your resources

There are some areas that need to be modified. The very first function, __Set-MyAzureEnvironment__, is where the Azure tenants are configured; add the tenant info to each site. Review Lines 64-75

```powershell
#My Azure Site A lab
'Resource Tenant' {
        $myTenantID = '<your tenant ID>'
        $mySubscriptionName = '<your subscription name>'
        $mySubscriptionID = '<your subscription ID>'
        $myResourceGroup = '<your resource group>'
    }
#My Azure Site B lab
'Services Tenant' {
        $myTenantID = '<your tenant ID>'
        $mySubscriptionName = '<your subscription name>'
        $mySubscriptionID = '<your subscription ID>'
        $myResourceGroup = '<your resource group>'
    }
```

Make sure to update _[ValidateSet()]_ in the param section (line 25) as well if you add new or change environments names.

```powershell
[ValidateSet('Resource Tenant','Services Tenant')]
```

I have the script installs and updates the required modules. It currently monitors these modules:

- Az
- Azure
- AzureAD
- Az.Security

I would recommend keeping the above

On line 29, you can edit it and add more if you like. :

```powershell
#set this to what you want
$Checkmodules = @('Az','Az.Security','Azure','AzureAD')
```

## Call the script

Copy this script to your __C:\Users\\\<userprofile>\Documents\WindowsPowerShell__ folder and relaunch Windows PowerShell.
> __Keep in Mind__: If you already have a profile script, make sure you make a backup or integrate your code into this one.

 I've added a check to see if the script is called directly and if it is, it exits the script; this resolves issues with services calling powershell or VS code calling the script  each time its launched

The only command you will need run is:

- Start-MyLabEnvironment

This will set off a chain of events

If you have not specified the correct Azure Tenant info, You will also be presented with a grid output to verify the subscription list.

Also the first time you connect to Azure using PowerShell, an identity file is creating in your profile. Being a nerd, the script parses that file to look for the authenticated username's first name and output a voice such as: "Good morning Dick, Please wait while I check for installed modules..."

However you can disable it if you set line 14 to

```powershell
$VoiceWelcomeMessage = $false

$DefaultVoiceProfile = 'Female'
```
there are othere global settings that can be changes for more automation and features (lines 18-24)

```powershell
$global:MyLabTag = 'StartupOrder'

$global:MyLabTenant = 'Resource Tenant'

$global:MyMDTSimulatorPath = 'C:\MDTSimulator'

$global:MyDeploymentShare = "\\$env:ComputerName\Deploymentshare$"
```

## Future changes
- I plan on making more voice commands within the functions; yes I know what your thinking...NERD ALERT!!

## Screenshots

This is what the startup looks like.
![Console](.images/AzureEnvironment.PNG)

NSG assigned to __subnet__ and only enabled JIT for VMs with public IP's
![NSG On Subnet](.images/status.png)

NSG assigned to VMs __nic__ and only enabled JIT for the ones with public IP's
![NSG on NIC](.images/startedvms.png)

## Functions Included

THe functions are the main functions to manage your virtual environment. However there are a lot of other functions available in this script:

UPDATE 06/11/2021: All command verbs will start with _My_


- **Start-MyLabEnvironment** --> Uses global variables start up lab. This will include starting Azure and Hyper-V labs and making sure Azure gateway IP is accurate. Only starts VM's with order tag specified in global variable. Set tag in notes of Hyper-V VM (eg: StartupOrder: 1)
- **Start-MyAzureEnvironment** --> Underlying function that will start the VM's in Azure environment
- Start-MyAzureVM
- Start-MyElevatedProcess
- Start-MyMDTSimulator
- Start-MyVSCodeInstall
- Start-MyHyperVM
- **Connect-MyAzureEnvironment** --> Starts just the Azure lab, 
- Enable-MyAzureJitPolicy
- Get-MyAzureNSGRules
- Get-MyAzureUserName
- Get-MyAzureVM
- Get-MyHyperVM
- Get-MyRandomAlphanumericString
- Get-MyRandomSerialNumber
- Get-MyRemoteDesktopData
- Get-MyVolumeLevel
- Install-MyLatestModule
- Open-MyFile
- Out-MyVoice
- Restart-MyHyperV
- Set-MyAzureEnvironment
- Set-MyAzureJitPolicy
- Set-MyVolumeLevel
- Set-MyWindowPosition
- Show-MyCommands
- Stop-MyHyperVM
- Test-MyIsAdmin
- Test-MyVSCode
- Test-MyVSCodeInstall
