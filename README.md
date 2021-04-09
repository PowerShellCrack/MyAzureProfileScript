# MyAzureProfileScript

THis is a still a work in progress...isn't everything?

Its a Profile script for your admin system to simply manage VMs in multiple (or singular) Azure tenants. I have multiple tenants and need to start them up based on what I am testing.

## Customize your resources

There are some areas that needs to be modified. The very first function, __Set-MyAzureEnvironment__, is where the Azure tenants are configured; add the tenant info to each site. Review Lines 38-51

```powershell
#My Azure Site B lab
'SiteB' {
            $global:myTenantID = '<your tenant ID>'
            $global:mySubscriptionName = '<your subscription name>'
            $global:mySubscriptionID = '<your subscription ID>'
            $global:myResourceGroup = '<your resource group>'
        }
```
Make sure to update [ValidateSet] in the param section (line 25) as well if you add new site names.

```powershell
[ValidateSet('SiteA','SiteB')]
```

On line 11, I have the script install the required modules. You can edit it and add more if you like. I would recommend keeping these (used by this script):

- Az
- Azure
- AzureAD
- Az.Security

## Call the script
Copy this script to your C:\Users\\\<userprofile>\Documents\WindowsPowerShell folder and relaunch Windows PowerShell.
> __Keep in Mind__: If you already have a profile script, make sure you make a backup or integrate your code into this one.

Since the script would be loaded in your profile, it would be called by anything that runs in with your context including corporate scripts that are managing your device. To fix this, I've added a check to see if the script is called directly and if it is, it exits the script

The only command you will need run is:
 - Start-MyAzureEnvironment

This will set off a chain of events
You will also be presented with a grid output to verify the subscription list.

Also the first time you connect to Azure using PowerShell, an identity file is creating in your profile. And me being a nerd, I parse that file to look for the authenticated username's first name and output a voice such as: "Good morning Dick,Please wait while I check for installed modules..."

However you can disable it if you set line 8 to
```powershell
$VoiceWelcomeMessage = $false
```

I plan on making more voice commands within the functions; yes I know what your thinking...NERD ALERT!!

## Screenshots
This is what the startup looks like.
![Console](.images/AzureEnvironment.PNG)

NSG assigned to __subnet__ and only enabled JIT for VMs with public IP's
![NSG On Subnet](.images/status.png)

NSG assigned to VMs __nic__ and only enabled JIT for the ones with public IP's
![NSG on NIC](.images/startedvms.png)

## Functions Included
THe functions are the main functions to manage your virtual environment. However there are a lot of other functions available in this script:

- Test-IsISE
- Test-VSCode
- Test-VSCodeInstall
- Start-VSCodeInstall
- Install-LatestModule
- Connect-MyAzureEnvironment
- Get-MyAzureNSGRules
- Get-MyAzureVM
- Set-MyJitAccess
- Set-MyAzureEnvironment
- Start-MyAzureVM
- Start-MyAzureEnvironment
- Get-MyAzureUserName
- Get-MyHyperVM
- Kill-MyHyperVM
- Restart-MyHyperV
- Show-MyCommands