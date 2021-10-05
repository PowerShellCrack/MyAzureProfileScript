#=======================================================
# VARIABLES
#=======================================================
#exit profile script if a script is called instead
if ([Environment]::GetCommandLineArgs().Count -gt 1) {
    #log whats going on, then exit
    [string]$FileName = 'Profile_RunningScripts' + '_' + (get-date -Format MM-dd-yyyy) + '.log'
    [Environment]::GetCommandLineArgs() | Out-File $env:TEMP\$FileName -Force -Append
    exit
}

#voice
$VoiceWelcomeMessage = $true

$DefaultVoiceProfile = 'Female'

$global:MyLabTag = 'StartupOrder'

$global:MyLabTenant = 'Resource Tenant'

$global:MyLabRouterIP = '192.168.21.6'

$global:MyPublicIP = Invoke-RestMethod 'http://ipinfo.io/json' | Select-Object -ExpandProperty IP

#set this to what you want
$Checkmodules = @('Az','Az.Security','Azure','AzureAD')

#=======================================================
# Functions
#=======================================================
Function Set-MyAzureEnvironment{
    [CmdletBinding()]
    param(
        [ValidateSet('Resource Tenant','Services Tenant')]
        [string]$MyEnv,

        [ValidateSet('TenantID','SubscriptionName','SubscriptionID','ResourceGroup')]
        [string]$Output,

        [switch]$Force,

        [boolean]$OutVoice = $VoiceWelcomeMessage
    )
    ## Get the name of this function
    [string]${CmdletName} = $MyInvocation.MyCommand
    #build log name
    [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
    Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

    #if parameter force is set, always show selection
    If($PSBoundParameters.ContainsKey('MyEnv') -and !$Force){
        $global:MyEnv = $MyEnv
    }
    Elseif ($Force -or (!$global:MyEnv -or !$global:MyTenantID -or !$global:MySubscriptionName -or !$global:MySubscriptionID -or !$global:MyResourceGroup) ) {
        If($OutVoice){Out-MyVoice "You must select an Azure Environment"}
        $global:MyEnv = Get-ParameterOption -Command ${CmdletName} -Parameter MyEnv | Out-GridView -Title "Select an Azure Environment" -PassThru
    }

    Switch($global:myEnv){
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
        #Default to Site A lab
        default {
            $myTenantID = '<your tenant ID>'
            $mySubscriptionName = '<your subscription name>'
            $mySubscriptionID = '<your subscription ID>'
            $myResourceGroup = '<your resource group>'
    }
    }

    Switch($Output){
        'TenantID' {$global:myTenantID = $myTenantID}
        'SubscriptionName' {$global:mySubscriptionName = $mySubscriptionName }
        'SubscriptionID' {$global:mySubscriptionID = $mySubscriptionID}
        'ResourceGroup' {$global:myResourceGroup = $myResourceGroup}

        default {
            $global:myTenantID = $myTenantID
            $global:mySubscriptionName = $mySubscriptionName
            $global:mySubscriptionID = $mySubscriptionID
            $global:myResourceGroup = $myResourceGroup
        }

    }

    Stop-Transcript | Out-Null
}

#region FUNCTION: Get parameter values from cmdlet
#https://michaellwest.blogspot.com/2013/03/get-validateset-or-enum-options-in_9.html
function Get-ParameterOption {
    param(
        $Command,
        $Parameter
    )

    $parameters = Get-Command -Name $Command | Select-Object -ExpandProperty Parameters
    $type = $parameters[$Parameter].ParameterType
    if($type.IsEnum) {
        [System.Enum]::GetNames($type)
    } else {
        $parameters[$Parameter].Attributes.ValidValues
    }
}
#endregion


#region FUNCTION: Check if running in Visual Studio Code
Function Test-MyVSCode{
    if($env:TERM_PROGRAM -eq 'vscode') {
        return $true;
    }
    Else{
        return $false;
    }
}
#endregion


Function Test-MyVSCodeInstall{
    $Paths = (Get-Item env:Path).Value.split(';')
    If($paths -like '*Microsoft VS Code*'){
        return $true
    }Else{
        return $false
    }
}

Function Start-MyVSCodeInstall {
    $uri = 'https://raw.githubusercontent.com/PowerShell/vscode-powershell/master/scripts/Install-VSCode.ps1'
    #Invoke-Command -ScriptBlock ([scriptblock]::Create((Invoke-WebRequest $uri -UseBasicParsing).Content))

    If(-Not(Test-MyVSCodeInstall) ){
        $Code = (Invoke-WebRequest $uri -UseBasicParsing).Content
        $code | Out-File $env:temp\vsc.ps1 -Force
        Unblock-File $env:temp\vsc.ps1
        . $env:temp\vsc.ps1 -LaunchWhenDone
    }Else{
        code
    }
}


#connect to volume controller on device
$VolumeController = @'
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume
{
    // f(), g(), ... are unused COM method slots. Define these if you care
    int f(); int g(); int h(); int i();
    int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
    int j();
    int GetMasterVolumeLevelScalar(out float pfLevel);
    int k(); int l(); int m(); int n();
    int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
    int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice
{
    int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator
{
    int f(); // Unused
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio
{
    static IAudioEndpointVolume Vol()
    {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
        IAudioEndpointVolume epv = null;
        var epvid = typeof(IAudioEndpointVolume).GUID;
        Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
        return epv;
    }
    public static float Volume
    {
        get { float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty)); }
    }
    public static bool Mute
    {
        get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
    }
}
'@

Function Get-MyVolumeLevel{

    Try{
        Add-Type -TypeDefinition $VolumeController | Out-Null
    }
    Catch{

    }
    Finally{
        If([audio]::Mute){
            0
        }
        Else{
            [int]([audio]::Volume * 100)
        }
    }
}


Function Set-MyVolumeLevel{
    param(
        [parameter(Mandatory=$true)]
        [ValidateRange(1,100)]
        [int]$Volume,
        [switch]$Mute
    )
    $SetVol = $Volume/100 # 0.1 = 10%, etc.

    Try{
        Add-Type -TypeDefinition $VolumeController | Out-Null
    }
    Catch{

    }
    Finally{
        If($mute){
            [audio]::Mute = $true
        }Else{
            [audio]::Mute = $false
            [audio]::Volume  = $SetVol
        }
    }
}

Function Out-MyVoice
{
    <#
    .SYNOPSIS
    Reads a message to the user.

    .DESCRIPTION
    Uses the default voice of the client to read a message
    to the user.

    .PARAMETER Message
    The string of text to read.

    .PARAMETER VoiceType
    Allows for the default choice to be changed using the
    default voices installed on Windows 10. Acceptable values are:
    Male
    Female

    .PARAMETER PassThru
    Passes the piped in object to the pipeline.

    .EXAMPLE
    Out-MyVoice "Script Complete"

    Reads back to the user "Script Complete"

    .EXAMPLE
    $CustomObject | Out-MyVoice -PassThru

    If the object has a property called "VoiceMessage" and is of
    data type [STRING], then the string is read.  If the object
    contains a "VoiceType" parameter that is valid, that
    voice will be used. The original object is then passed
    into the pipeline.

    .EXAMPLE
    Get-WmiObject Win32_Product |
    ForEach -process {Write-Output $_} -end{Out-MyVoice -VoiceMessage "Script Completed"}

    Recovers the product information from WMI and the notifies the
    user with the voice message "Script Completed" while also
    passing the results to the pipeline.

    .EXAMPLE
    Start-Job -ScriptBlock {Get-WmiObject WIn32_Product} -Name GetProducts
    While ((Get-job -Name GetProducts).State -ne "Completed")
    {
        Start-sleep -Milliseconds 500
    }
    Out-MyVoice -VoiceMessage "Done"

    Notifies the user when a background job has completed.

    .NOTES
    Tested on Windows 10
    #>
    [cmdletBinding()]
    Param
    (
        [parameter(ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True)]
                   [String]$VoiceMessage,
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [ValidateSet("Male","Female")]
                   [String]$VoiceType = $DefaultVoiceProfile,
                   [Switch]$PassThru,
        [parameter(ValueFromPipeline=$True)]
                   [PSObject]$InputObject
    )
    BEGIN
    {
        # When the cmdlet starts, create a new
        # SAPI voice object.
        $voice = New-Object -com SAPI.SpVoice

        # If the client is Windows 8, then allow for different voices.
        If ((Get-CimInstance -ClassName Win32_Operatingsystem).Name -Like "*Windows 1*")
        {
            # Get a list of all voices.
            $Voice.GetVoices() | Out-Null
            $voices = $Voice.GetVoices();
            $V = @()
            ForEach ($Item in $Voices)
            {
                $V += $Item
            }
            # Set the voice to use using the $VoiceType parameter.
            # The default voice will be used otherwise.
            Switch ($VoiceType)
            {
                "Male" {$Voice.Voice = $V[0]}
                "Female" {$Voice.Voice = $V[1]}
            }
        } # End: IF Statement.
    }

    PROCESS
    {
        # If an array of messages is passed, this will allow
        # for each message to be read.
        ForEach ($M in $VoiceMessage)
        {
            # Speak the message.
            $voice.Speak($M) | Out-Null
        }
    } # End: ForEach ($M in $VoiceMessage)
    END
    {
        If ($PassThru)
        {
            Write-Output $InputObject
        }
    } #End: PROCESS
} # End: Out-MyVoice


Function Install-MyLatestModule {
    [CmdletBinding(DefaultParameterSetName = 'NameParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'NameParameterSet')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Name,

        [Parameter()]
        [switch]
        $Force,

        [ValidateSet("Always","Daily","Weekly","Monthly")]
        [string]$Frequency,

        [Parameter()]
        [switch]
        $AllowImport
    )

    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        if (-not $PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
        }

        if (-not $PSBoundParameters.ContainsKey('WhatIf')) {
            $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference')
        }

        [string]$ModuleName = $null
        $LatestModule = $null
        $ExistingModules = $null

        Write-Host ("{0} :: Checking for latest installed modules [Press Ctrl+C to cancel]..." -f ${CmdletName})  -ForegroundColor Gray

        $tagOrderNumdate = (Get-date -Format "yyyy-MM-dd")

        #grab all version of the module installed
        $InstalledModules = Get-InstalledModule -ErrorAction SilentlyContinue
    }
    Process{
        Try{
            $psrepos = Get-PSRepository -ErrorAction Stop

            foreach($repo in $psrepos){
                $RepoName = $repo.Name
                If($repo.InstallationPolicy -eq 'Untrusted'){
                    Set-PSRepository -Name $RepoName -InstallationPolicy Trusted
                }

            }
        }
        Catch{
            $_.exception.message
            break
        }

        #Import csv file first
        If(Test-path $env:USERPROFILE\.modulecheck -ErrorAction SilentlyContinue){
            $DateChecked = Import-Csv $env:USERPROFILE\.modulecheck -ErrorAction SilentlyContinue
        }Else{
            $DateChecked = $null
            #build a fresh csv file
            Set-Content -Value "ModuleName,DateChecked" -Path $env:USERPROFILE\.modulecheck
        }

        #remove-item $env:USERPROFILE\.modulecheck -force
        #$testdate = [DateTime]::Today.AddDays(-31).ToString("yyyy-MM-dd")
        #Add-Content -Value "Az, $testdate" -Path $env:USERPROFILE\.modulecheck
        #Add-Content -Value "Az.Security, $testdate" -Path $env:USERPROFILE\.modulecheck
        #TEST $item = $Checkmodules[0]
        #TEST $item = 'Az'



        foreach ($item in $name)
        {
            $ModuleLastRunDate = $DateChecked | Where-Object ModuleName -eq $item | Select-Object -ExpandProperty DateChecked -Last 1
            If($ModuleLastRunDate)
            {
                switch($Frequency){
                    "Always"  {$CheckModule = $true;$CheckMsg = "validating version"}
                    "Daily"   {$CheckModule = (Get-date $ModuleLastRunDate) -ne $tagOrderNumdate;$CheckMsg = "skipped validation until tomorrow"}
                    "Weekly"  {$CheckModule = (Get-date $ModuleLastRunDate) -le [DateTime]::Today.AddDays(-7).ToString("yyyy-MM-dd");$CheckMsg = "skipped validation until next week"}
                    "Monthly" {$CheckModule = (Get-date $ModuleLastRunDate) -le [DateTime]::Today.AddDays(-30).ToString("yyyy-MM-dd");$CheckMsg = "skipped validation until next month"}
                    default   {$CheckModule = $true;$CheckMsg = "validating version"}
                }
            }
            Else{
                $CheckModule = $true
            }

            #format out text
            Write-Host ("Searching for Module: ") -ForegroundColor Gray -NoNewline
            Write-Host ("{0}" -f $item) -ForegroundColor White -NoNewline
            Write-Host ("...") -ForegroundColor Gray -NoNewline

            [string]$ModuleName = $item

            $ExistingModules = $InstalledModules | Where-Object Name -eq $ModuleName

            #comment on module identified
            If($ExistingModules){
                Write-Host ("found") -ForegroundColor green -NoNewline
            }Else{
                Write-Host ("not found") -ForegroundColor red -NoNewline
            }
            Write-Host ("...") -ForegroundColor Gray -NoNewline

            #if scheduled to check module, search for module online
            If($CheckModule)
            {
                If($ExistingModules){
                    Write-Host ("Checking if version [{0}] is latest..." -f $ExistingModules.Version.ToString()) -ForegroundColor Yellow
                }

                $LatestModule = Find-Module $ModuleName -ErrorAction SilentlyContinue
            }
            Else{
                 Write-Host ("{0}" -f $CheckMsg) -ForegroundColor yellow
                $LatestModule = $ExistingModules
                Continue
            }

            #if latest module has been found online, proceed
            If($null -ne $LatestModule)
            {

                #ignore any versions installed, uninstall all and install latest
                Try
                {
                    If($PSBoundParameters.ContainsKey('Force'))
                    {
                        Write-Host ("re-installing module [{0}]..." -f $ModuleName) -ForegroundColor Cyan -NoNewline
                        $ExistingModules | Uninstall-Module -Force -ErrorAction Stop
                        Install-Module $ModuleName -RequiredVersion $LatestModule.Version -Scope AllUsers -Force -SkipPublisherCheck -ErrorAction Stop -Verbose:$VerbosePreference
                        Write-Host ("Completed") -ForegroundColor Green
                    }
                    Else
                    {
                        #if no moduels exist
                        If($null -eq $ExistingModules)
                        {
                            Write-Host ("[{0}] is not installed, installing..." -f $ModuleName) -ForegroundColor Gray -NoNewline
                            Install-Module $ModuleName -Scope AllUsers -Force -SkipPublisherCheck  -AllowClobber -ErrorAction Stop -Verbose:$VerbosePreference
                            Write-Host ("Installed") -ForegroundColor Green
                        }

                        #are there multiple of the same module installed?
                        ElseIf( ($ExistingModules | Measure-Object).Count -gt 1)
                        {
                            Write-Host ("multiple modules found named [{0}], cleaning..." -f $ModuleName) -ForegroundColor Gray -NoNewline

                            If($LatestModule.Version -in $ExistingModules.Version)
                            {
                                Write-Host ("Cleaning up older [{0}] modules..." -f $ModuleName) -ForegroundColor Yellow -NoNewline
                                #Check to see if latest module is installed already and uninstall anything older
                                $ExistingModules | Where-Object Version -NotMatch $LatestModule.Version | Uninstall-Module -Force -ErrorAction Stop
                            }
                            Else
                            {
                                #uninstall all older Modules with that name, then install the latest
                                Write-Host ("Uninstalling older [{0}] modules and installing the latest module for [{0}]..." -f $ModuleName) -ForegroundColor Yellow -NoNewline
                                Get-Module -FullyQualifiedName $ModuleName -ListAvailable | Uninstall-Module -Force -ErrorAction Stop
                                Install-Module $ModuleName -RequiredVersion $LatestModule.Version -Scope AllUsers -AllowClobber -Force -SkipPublisherCheck -ErrorAction Stop -Verbose:$VerbosePreference
                            }
                            Write-Host ("done") -ForegroundColor Green
                        }

                        #if only one module exist but not the latest version
                        ElseIf($ExistingModules.Version -ne $LatestModule.Version)
                        {
                            Write-Host ("found newer version [{0}]..." -f $LatestModule.Version.ToString()) -ForegroundColor yellow -NoNewline
                            #Update module since it was found
                            If($VerbosePreference){Write-Host ("Updating Module [{0}] from [{1}] to the latest version [{2}]..." -f $ModuleName,$ExistingModules.Version,$LatestModule.Version) -NoNewline -ForegroundColor Yellow}
                            Update-Module $ModuleName -RequiredVersion $LatestModule.Version -Force -ErrorAction Stop -Verbose:$VerbosePreference
                            Write-Host ("Updated") -ForegroundColor Green
                        }
                        Else
                        {
                            #No issue
                            Write-Host ("Module [{0}] with version [{1}] is the latest!" -f $ModuleName,$ExistingModules.Version) -ForegroundColor Green
                            Continue
                        }
                    }
                }
                Catch
                {
                    Write-Host ("Failed. Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
                }
                Finally
                {
                    If($AllowImport){
                        #importing module
                        Write-Host ("Importing Module [{0}] for use..." -f $ModuleName) -ForegroundColor Green
                        Import-Module -Name $ModuleName -Force:$force -Verbose:$VerbosePreference
                    }

                    #set module and date for today
                    Add-Content -Value "$ModuleName, $tagOrderNumdate" -Path $env:USERPROFILE\.modulecheck
                }
            }
            Else{
                If($VerbosePreference){Write-Host ("Module [{0}] does not exist, unable to update" -f $ModuleName) -ForegroundColor Red}
            }

        } #end of module loop
    }
    End{
        If($VerbosePreference){Write-Host ("{0} :: Completed module check" -f ${CmdletName}) -ForegroundColor Gray}
        Stop-Transcript | Out-Null
    }
}


Function Connect-MyAzureEnvironment{
    [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $false,Position = 0)]
        [string]$TenantID,

        [Parameter(Mandatory = $false,
            Position = 1,
            ParameterSetName = 'NameParameterSet')]
        [string]$SubscriptionName,

        [Parameter(Mandatory = $false,
            Position = 1,
            ParameterSetName = 'IDParameterSet')]
        [string]$SubscriptionID ,

        [Parameter(Mandatory = $false,
        ValueFromPipelineByPropertyName = $true,
            Position = 2)]
        [string]$ResourceGroupName,

        [switch]$ClearAll,

        [boolean]$OutVoice = $VoiceWelcomeMessage
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        if (-not $PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
        }

        #grab global variables
        Set-MyAzureEnvironment

        #overwrite global variable if specified
        if ($PSBoundParameters.ContainsKey('TenantID')) {
            $global:MyTenantID = $TenantID
        }

        if ($PSBoundParameters.ContainsKey('SubscriptionID')) {
            $global:MySubscriptionID = $SubscriptionID
        }

        if ($PSBoundParameters.ContainsKey('SubscriptionName')) {
            $global:MySubscriptionName = $SubscriptionName
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
            $global:MyResourceGroup = $ResourceGroupName
        }

        If($ClearAll){
            Write-host ("There was an issue signing into Azure. Resetting and attempting again...") -ForegroundColor yellow
            Clear-AzDefault -ErrorAction SilentlyContinue -Force
            Clear-AzContext -ErrorAction SilentlyContinue -Force
            Disconnect-AzAccount -ErrorAction SilentlyContinue
        }
        Else{
            Try{
                #grab current AZ resources
                $Context = Get-AzContext -ErrorAction Stop
                #if defualt is not set, attempt to set it
                If($null -eq (Get-AzDefault) )
                {
                    $DefaultRG = Set-AzDefault -ResourceGroupName $global:MyResourceGroup -Force
                }Else{
                    $DefaultRG = Get-AzDefault -ErrorAction Stop
                }
            }Catch{
                Write-host ("Failed to get Azure context. {0}" -f $_.Exception.Message) -ForegroundColor yellow
            }
        }

        $MySubscriptions = @()
        $MyRGs = @()

        If($VerbosePreference){Write-Host ''}
    }
    Process{
        If($VerbosePreference){Write-Host ("Attempting to connect to Azure...") -ForegroundColor Yellow -NoNewline}
        #region connect to Azure if not already connected
        Try{
            If(($null -eq $Context.Subscription.SubscriptionId) -or ($null -eq $Context.Subscription.Name))
            {
                If($global:MyTenantID){
                    $AzAccount = Connect-AzAccount -Tenant $global:MyTenantID -ErrorAction Stop
                }Else{
                    $AzAccount = Connect-AzAccount -ErrorAction Stop
                }
                If($OutVoice){Out-MyVoice "You must select an Azure Subscription"}
                $AzSubscription += Get-AzSubscription -WarningAction SilentlyContinue | Out-GridView -PassThru -Title "Select a valid Azure Subscription" | Select-AzSubscription -WarningAction SilentlyContinue
                Set-AzContext -Tenant $AzSubscription.Subscription.TenantId -Subscription $AzSubscription.Subscription.id | Out-Null
                If($VerbosePreference){Write-Host ("Successfully connected to Azure!") -ForegroundColor Green}

                $MyRGs += Get-AzResourceGroup | Select-Object -ExpandProperty ResourceGroupName
                <#
                If(($global:MyResourceGroup -notin $MyRGs) -or ($DefaultRG.Name -ne $global:MyResourceGroup)){
                    $global:MyResourceGroup = Get-AzResourceGroup | Out-GridView -PassThru -Title "Select a Azure Resource Group" | Select-Object -ExpandProperty ResourceGroupName
                    #set the new context based on found RG
                    Set-AzDefault -ResourceGroupName $global:MyResourceGroup -Force | Out-Null
                }
                #>
            }
            Else{
                If($VerbosePreference){Write-Host ("Already connected to Azure using account [{0}] and subscription [{1}]" -f $Context.Account.Id,$Context.Subscription.Name) -ForegroundColor Green}
            }
        }
        Catch{
            If($VerbosePreference){Write-Host ("Unable to connect to Azure account with credentials: {0}. Error: {1}" -f $AzAccount.Context.Account.Id, $_.Exception.Message) -ForegroundColor Red}
            Break
        }
        Finally{
            #To suppress these warning messages
            Write-Verbose ("Suppressing Azure Powershell change warnings...")
            Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true" | Out-Null

            #grab all resource groups in Azure
            $MyRGs += Get-AzResourceGroup | Select-Object -ExpandProperty ResourceGroupName
            #Determine if azure resoruce group is in the same list as script
            If( ($global:MyResourceGroup -notin $MyRGs) -or ($DefaultRG.Name -ne $global:MyResourceGroup) )
            {
                If($OutVoice){Out-MyVoice "You must select an Azure Resource Group"}
                $global:MyResourceGroup = Get-AzResourceGroup | Out-GridView -PassThru -Title "Select a Azure Resource Group" | Select-Object -ExpandProperty ResourceGroupName
                #set the new context based on found RG
                Set-AzDefault -ResourceGroupName $global:MyResourceGroup -Force | Out-Null
            }

            #set the global values if connection
            If($AzSubscription){
                $global:MySubscriptionName = $AzSubscription.Subscription.Name;
                $global:MySubscriptionID = $AzSubscription.Subscription.Id;
                $global:MyTenantID = $AzSubscription.Subscription.TenantId;
            }
        }
    }
    End{
        #once logged in, set defaults context
        Get-AzContext
        Stop-Transcript | Out-Null
    }

}


Function Get-MyAzureNSGRules{
    [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [Alias("VM")]
        [string[]]$VMname,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [Alias("ResourceGroup")]
        [string]$ResourceGroupName
    )
    Begin
    {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #attempt connection to Azure. If connection is not successfull clear all settings and try again
        Try{
            $Null = Connect-MyAzureEnvironment -ErrorAction Stop
        }
        Catch{
            Connect-MyAzureEnvironment -ClearAll
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
             $global:MyResourceGroup = $ResourceGroupName
        }

        #if resource group is blank, throw an error
        If($null -eq $ResourceGroupName){
            Throw "Resource group name has not been specified"
        }

        $nics = @()
        $report = @()
        $subnetIds = @()

        $Vnets = Get-AzVirtualNetwork
        $NSGs = Get-AzNetworkSecurityGroup
    }
    Process
    {
        #get subnets
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            #TEST $VM = $VMname
            Foreach($VM in $VMName)
            {
                If(Get-MyAzureVM -VMname $VM){
                    $nics += Get-AzNetworkInterface | Where-Object { $_.VirtualMachine.id -like "*$VM*"}
                    #$nics += Get-AzNetworkInterface | Where-Object { $_.VirtualMachine.Id.tostring().substring($_.VirtualMachine.Id.tostring().lastindexof('/')+1) -eq $VM}
                    $subnetIds += $nics.IpConfigurations.subnet.id
                }
            }
        }
        Else{
            $nics = Get-AzNetworkInterface | Where-Object {$null -ne $_.VirtualMachine} #skip Nics with no VM
            $subnetIds = $nics.IpConfigurations.subnet.id
        }

        #get only unique subnets
        $subnetIds = $subnetIds | Select-Object -Unique

        #loop through all subnets
         #TEST $SubnetId = $subnetIds
        Foreach($SubnetId in $subnetIds)
        {
            $subnetName = ($subnetId.Split("/")[-1])
            #grab nsg name from vnet
            $VnetSubnet = $Vnets | Where-Object {$_.Subnets.Id -eq $subnetId}
            #check if NSG exists on subnet
            $VnetNSGId = $VnetSubnet.Subnets.NetworkSecurityGroup.id
            #$Null -ne $NSGs.NetworkInterfaces
            #$Null -ne $NSGs.Subnets
            #if NSG exists on subnet; grab the name and configs
            If($null -ne $VnetNSGId){
                $VnetNSGName = ($VnetNSGId).Split("/")[-1]
                $NSGConfigs = $NSGs | Where-Object Name -eq $VnetNSGName
            }
            #if NSG exists on elsewhere; grab the name and configs
            ElseIf($NSGs.SecurityRules.count -gt 1){
                $NSGConfigs = $NSGs
            }
            Else{
                #Write-Host "No NSG's were found." -ForegroundColor Red
                $report = 'disabled'
                Continue
            }

            Foreach($NSG in $NSGConfigs)
            {
                Foreach($rule in $NSG.SecurityRules){
                    #build info object
                    $info = "" | Select-Object NSGName,RuleName,Description,AttachedSubnet,Protocol,SourcePort,DestinationPort,SourceAddress,DestinationAddress,Access
                    $info.NSGName = $NSG.Name
                    $info.RuleName = $rule.Name
                    $info.Description = $rule.Description
                    $info.Protocol = $rule.Protocol
                    $info.AttachedSubnet = $SubnetName
                    $info.SourcePort = $rule.SourcePortRange | Foreach-Object {$_ -join ","}
                    $info.DestinationPort = $rule.DestinationPortRange | Foreach-Object {$_ -join ","}
                    $info.SourceAddress = $rule.SourceAddressPrefix | Foreach-Object {$_ -join ","}
                    $info.DestinationAddress = $rule.DestinationAddressPrefix | Foreach-Object {$_ -join ","}
                    $info.Access = $rule.Access
                    $report+=$info
                }
            }
        }
    }
    End
    {
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            Foreach($nic in $nics)
            {
                $report | Where-Object {$_.DestinationAddress -eq $nic.IpConfigurations.PrivateIpAddress}
            }
        }
        Else{
            $report
        }
        Stop-Transcript | Out-Null
    }
}

Function Get-MyAzureVM{
    [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [ArgumentCompleter( {
            param ( $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )

            $Global:MyAzVMs.VMname | Where-Object {
                $_ -like "$wordToComplete*"
            }

        } )]
        [Alias("VM")]
        [string]$VMname,

        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true)]
        [Alias("ResourceGroup")]
        [string]$ResourceGroupName = $global:MyResourceGroup,

        [switch]$NetworkDetails,

        [switch]$NoStatus
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #attempt connection to Azure. If connection is not successfull clear all settings and try again
        Try{
            $Null = Connect-MyAzureEnvironment -ErrorAction Stop
        }
        Catch{
            Connect-MyAzureEnvironment -ClearAll
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
             $global:MyResourceGroup = $ResourceGroupName
        }

        #if resource group is blank, throw an error
        If($null -eq $ResourceGroupName){
            Throw "Resource group name has not been specified"
        }

        #default to showing status (like verbose) unless noStatus is in param
        if ($PSBoundParameters.ContainsKey('NoStatus')) {
            $ShowStatus=$false
        }
        else{
            $ShowStatus=$true
        }

        if ($PSBoundParameters.ContainsKey('NetworkDetails'))
        {
            If($ShowStatus){Write-Host 'Collecting network information...' -ForegroundColor DarkGray -NoNewline}
            #grab all public IP's
            $PublicIPs = Get-AzPublicIpAddress

            $NSGRules = Get-MyAzureNSGRules
            If($ShowStatus){Write-Host 'completed' -ForegroundColor Green}
        }

        #show the appropiate message
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            If($ShowStatus){Write-Host ("Collecting Azure VM with names [{0}]..." -f ($VMName -join ",")) -ForegroundColor DarkGray -NoNewline}
        }Else{
            If($ShowStatus){Write-Host "Collecting all Azure VM's..." -ForegroundColor DarkGray -NoNewline}
        }

        $vmNics = @()
        $VMs = @()
        $report = @()

        #pull all NIC that are attached to Virtual Machine's
        $nics = Get-AzNetworkInterface | Where-Object {$null -ne $_.VirtualMachine}
        #pull all Virtual Machine's; this is easier and faster than build VM's list one at a time.
        $AllVMs = Get-AzVM -Status -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
    Process{
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            #filter out only listed VM's
            #TEST $VM = $VMname
            Foreach($VM in $VMName)
            {
                $vmNics += $nics | Where-Object { $_.VirtualMachine.id -like "*$VM"}
                $VMs += $AllVMs | Where-Object {$_.Name -eq $VM}
            }
            If($ShowStatus){Write-Host 'completed' -ForegroundColor Green}
        }
        Else{
            $VMs = $AllVMs
            $vmNics = $nics
            If($ShowStatus){Write-Host 'completed' -ForegroundColor Green}
        }

        if ($PSBoundParameters.ContainsKey('NetworkDetails'))
        {
            #list VM and their network info
            #TEST $nic = $vmNics[4]
            foreach($nic in $vmNics)
            {
                If($ShowStatus){Write-Host '.' -ForegroundColor DarkGray -NoNewline}
                # add vm values to
                [psobject]$VM = $VMs | where-object -Property Id -eq $nic.VirtualMachine.id

                If($VM){
                    #get public access information
                    $pipinfo = ($PublicIPs | Where-Object { $_.IpConfiguration.Id -match "^$($nic.id)/" })

                    #filter NSG rules
                    If($NSGRules -ne 'disabled'){
                        $JITRules = $NSGRules | Where-Object {($_.DestinationAddress -eq $nic.IpConfigurations.PrivateIpAddress) -and ($_.SourceAddress -eq $global:MyPublicIP) -and ($_.Access -eq 'Allow')}
                        If($JITRules){$JITAccess = $true}Else{$JITAccess = $false}
                    }
                    Else{
                        $JITAccess = 'disabled'
                    }

                    If($VM.PowerState -eq 'vm running'){$vmstate = 'Running'}Else{$vmstate = 'Stopped'}

                    #build info object
                    $info = "" | Select-Object Id,VMName,ResourceGroup,HostName,Location,LocalIP,PublicIP,PublicDNS,State,Tags,JITAccess

                    $info.Id = $VM.id
                    $info.VMName = $VM.Name
                    $info.HostName = $VM.OSProfile.ComputerName
                    $info.ResourceGroup = $ResourceGroupName
                    $info.Location = $VM.Location
                    $info.LocalIP = $nic.IpConfigurations.PrivateIpAddress
                    $info.PublicIP = $pipinfo.IpAddress
                    $info.PublicDNS = $pipinfo.DnsSettings.Fqdn
                    $info.State = $vmstate
                    $info.Tags = $VM.Tags
                    $info.JITAccess = $JITAccess
                    $report+=$info
                }
            }

            return $report
        }
        Else{
            return $VMs
        }
    }
    End{Stop-Transcript | Out-Null}

}


Function Start-MyAzureVM{
    # Microsoft - Compute Resources
    #-----------------------------
   [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [ArgumentCompleter( {
            param ( $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )

            $Global:MyAzVMs | Where-Object {$_.State -ne 'Running'} | Select-Object -ExpandProperty VMName | Where-Object {
                $_ -like "$wordToComplete*"
            }

        } )]
        [Alias("VM")]
        [string]$VMname,

        [Parameter(Mandatory = $False,
            ValueFromPipelineByPropertyName = $true)]
        [Alias("ResourceGroup")]
        [string]$ResourceGroupName = $global:MyResourceGroup,
        [string]$OrderTag,
        [switch]$NoStatus
    )
    Begin{
         ## Get the name of this function
         [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

         #build log name
         [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
         Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #attempt connection to Azure. If connection is not successfull clear all settings and try again
        Try{
            $Null = Connect-MyAzureEnvironment -ErrorAction Stop
        }
        Catch{
            Connect-MyAzureEnvironment -ClearAll
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
             $global:MyResourceGroup = $ResourceGroupName
        }

        #if resource group is blank, throw an error
        If($null -eq $ResourceGroupName){
            Throw "Resource group name has not been specified"
        }

        #default to showing status (like verbose) unless noStatus is in param
        if ($PSBoundParameters.ContainsKey('NoStatus')) {
            $ShowStatus=$false
        }
        else{
            $ShowStatus=$true
        }

        $VMs = @()
    }
    Process{

        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
             #TEST $VM = $VMname
            Foreach($VM in $VMName)
            {
                $startedmessage = ("{0} is started already!" -f $VM)
                $VMs += Get-MyAzureVM -VMName $VM -NetworkDetails -NoStatus | Where-Object {$_.State -ne 'Running'}
            }
        }
        Else{
            $startedmessage = "All VM's are started already!"
            $VMs = Get-MyAzureVM -NetworkDetails -NoStatus | Where-Object {$_.State -ne 'Running'}
        }

        If($ShowStatus){Write-Host 'Collecting VM running status...' -ForegroundColor DarkGray -NoNewline}

        If($VMs.count -eq 0)
        {
            If($ShowStatus){Write-Host $startedmessage -ForegroundColor Green}
        }
        Else{
            If( $PSBoundParameters.ContainsKey('OrderTag') ){
                # Get the StartupOrder tag, if missing set to be run last (10)
                $taggedVMs = @{}

                #TEST $OrderTag = 'StartupOrder'
                # $VMs | Select-Object VMname,Tags
                #first grab vms with tags
                ForEach ($vm in $VMs) {
                    #$startupValue = $null
                    if ($vm.Tags[$OrderTag] -eq 0)
                    {
                        Continue
                        #do not add vm to list
                        #$startupValue = $vm.Tags['StartupOrder']
                    }
                    Else
                    {
                        If($VerbosePreference){write-host ('Found {0} tag for {1}; value is: {2}' -f $OrderTag,$vm.VMName,$vm.Tags[$OrderTag])}
                        $taggedVMs.Add($vm.VMName,$vm.Tags[$OrderTag])
                    }

                    #$taggedVMs.Add($vm.name,$startupValue)
                }


                #Count must be more that 0 to continue
                If($taggedVMs.count -gt 0)
                {
                    #get max value of startup count and make that the start if the next loop
                    $StartupEndCount = $taggedVMs.values | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum


                    #increment number to tag that are null
                    foreach($key in $taggedVMs.Keys.Clone())
                    {
                        If ($null -eq $taggedVMs[$key])
                        {
                            $StartupEndCount = $StartupEndCount + 1
                            If($VerbosePreference){write-host ('Adding tag: {0}={1}' -f $key,$StartupEndCount)}
                            $taggedVMs[$key] += $StartupEndCount
                        }
                    }

                    Write-Host ("[{0}] VM's tagged to start" -f $taggedVMs.count) -ForegroundColor Green

                    # Start in order from lowest vm
                    $tagOrderNum = $taggedVMs.values | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
                    $i = 0
                    #TEST ITERATION: $tagOrderNum = 2
                    Do{
                        #increment iteration count
                        $i++
                        #Always null VM to start
                        $tobeStarted = $null
                        # Get the VM tag that matched current iteration
                        $tobeStarted = $taggedVMs.GetEnumerator().Where({$_.Value -eq $tagOrderNum}) | Select-Object -ExpandProperty Key
                        If($tobeStarted)
                        {
                            #Grab resource id from VM
                            $VMResourceID = $VMs | Where-Object VMName -eq $tobeStarted | Select-Object -ExpandProperty ID

                            Write-Host ("  Attempting to start VM in order: {0}" -f $tobeStarted)
                            Start-AzVM -Id $VMResourceID -ResourceGroupName $ResourceGroupName -asJob | Out-Null
                            #Start-AzVM -id $VMResourceID -AsJob
                        }
                        Else{
                            Write-Host ("  No VM found with {0}: {1}" -f $OrderTag,$tagOrderNum)
                        }
                        #increment tag order
                        $tagOrderNum++
                    }
                    Until ($i -gt $taggedVMs.count)

                }
                Else{
                    Write-Host ("No running VM's were tagged to start") -ForegroundColor Yellow
                }

            }
            Else{
                #start all deallocated VM's
                foreach ($VM in $VMs)
                {
                    If($ShowStatus){write-host ("Attempting to start VM: {0}" -f $VM.VMName)}
                    Start-AzVM -Name $VM.VMName -ResourceGroupName $ResourceGroupName -asJob | Out-Null
                }
            }
        }
    }
    End{
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            Foreach($VM in $VMName)
            {
                $global:MyAzVMs = Get-MyAzureVM -VMName $VM -NetworkDetails -NoStatus
            }
        }
        Else{
            $global:MyAzVMs = Get-MyAzureVM -NetworkDetails -NoStatus
        }

        If($ShowStatus){
            $global:MyAzVMs | Select-Object VMName,HostName,LocalIP,PublicIP,PublicDNS,State,JITAccess | Format-Table
        }
        Stop-Transcript | Out-Null
    }
}


Function Get-MaxDuration ([string]$InStr) {
    $Out = $InStr -replace ("[^\d]")
    try {return [int]$Out}
    catch {}
    try {return [uint64]$Out}
    catch {return 0}
}

#https://github.com/CharbelNemnom/Power-MVP-Elite/blob/master/Request%20JIT%20VM%20Access/Request-JITVMAccess.ps1
#https://docs.microsoft.com/en-us/powershell/module/az.security/start-azjitnetworkaccesspolicy?view=azps-5.9.0
Function Set-MyAzureJitPolicy{
    [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [ArgumentCompleter( {
            param ( $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )

            $RemoteVMs = @()
            $RemoteVMs += $Global:MyAzVMs | Where-Object {$Null -ne $_.PublicDNS} | Select-Object -ExpandProperty PublicDNS
            $RemoteVMs += $Global:MyAzVMs | Where-Object {$Null -ne $_.PublicIP} | Select-Object -ExpandProperty PublicIP

            $RemoteVMs | Where-Object {
                $_ -like "$wordToComplete*"
            }

        } )]
        [Alias("VM")]
        [string]$VMname,

        [Parameter(Mandatory = $False,
            ValueFromPipelineByPropertyName = $true)]
        [Alias("ResourceGroup")]
        [string]$ResourceGroupName = $global:MyResourceGroup,

        [int]$Port = 3389,
        [int]$MaxTime = 3,
        [string]$SourceIP,

        [switch]$force,

        [switch]$NoStatus
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #attempt connection to Azure. If connection is not successfull clear all settings and try again
        Try{
            $Null = Connect-MyAzureEnvironment -ErrorAction Stop
        }
        Catch{
            Connect-MyAzureEnvironment -ClearAll
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
            $global:MyResourceGroup = $ResourceGroupName
        }

        #if resource group is blank, throw an error
        If($null -eq $ResourceGroupName){
            Throw "Resource group name has not been specified"
        }

        If(-Not($PSBoundParameters.ContainsKey('SourceIP')) ){
            $SourceIP = Invoke-RestMethod 'http://ipinfo.io/json' -Verbose:$false | Select-Object -ExpandProperty IP
        }

        #default to showing status (like verbose) unless noStatus is in param
        if ($PSBoundParameters.ContainsKey('NoStatus')) {
            $ShowStatus=$false
        }
        else{
            $ShowStatus=$true
        }

        #grab all current JIT requests first
        $VMJitAccessPolicy = Get-AzJitNetworkAccessPolicy
        #grab all NSG rules
        $NSGRules = Get-MyAzureNSGRules
        $VMs = @();
    }
    Process{
        If(!$VMJitAccessPolicy -or !$NSGRules -or ($NSGRules -eq 'disabled')){Continue}

        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            Foreach($VM in $VMName)
            {
                $VMs += Get-MyAzureVM -VMname $VM -NetworkDetails -NoStatus
            }
        }
        Else{
            $VMs = Get-MyAzureVM -NetworkDetails -NoStatus
        }

        foreach ($VM in $VMs)
        {
            $VMAllPortsDetails=@()

            If($VMJitAccessPolicy){
                $VMAllPortsDetails += $VMJitAccessPolicy.VirtualMachines | Where-Object { $_.Id.tostring().substring($_.Id.tostring().lastindexof('/')+1) -in $VM.VMName} | Select-Object -ExpandProperty Ports
                $VMAccessPorts = $VMAllPortsDetails | Select-Object -ExpandProperty Number -Unique
            }

            If($Port -notin $VMAccessPorts){
                If($ShowStatus){write-host ("JIT Policy does not allow port [{0}] for remote access on [{1}]" -f $Port,$VM.VMName) -ForegroundColor red}
                Continue  #stop current loop but CONTINUE next one
            }

            #$MaxTime = Get-MaxDuration ($VMAllPortsDetails.MaxRequestAccessDuration | Select-Object -First 1)
            $Date = (Get-Date).ToUniversalTime().AddHours($MaxTime)
            $endTimeUtc = Get-Date -Date $Date -Format o

            If($null -eq $SourceIP){
                If($ShowStatus){write-host ("Unable to get your Source IP. JIT request is canceled") -ForegroundColor red}
                Break
            }

            If($ShowStatus){Write-Host ("Validating Just-In-Time access for [{0}] on port [{1}]..." -f $VM.VMName,$Port) -NoNewline}

            $JITexists = $NSGRules | Where-Object { ($_.DestinationAddress -eq $VM.LocalIP) -and ($_.SourceAddress -eq $SourceIP) -and ($_.Access -eq 'Allow')}

            If(!$JITexists -or $force)
            {
                #build JIT policy
                $JitPolicy = (@{
                        id    = "$($VM.Id)"
                        ports = (@{
                                number                     = $Port;
                                endTimeUtc                 = "$endTimeUtc";
                                allowedSourceAddressPrefix = @("$SourceIP")
                            })

                    })
                $JitPolicyArr = @($JitPolicy)

                Try{
                    If($ShowStatus){Write-Host 'Requesting Just-in-time access...' -ForegroundColor Gray -NoNewline}
                    #$ResourceId = '/subscriptions/cb673656-d089-4158-a27a-628bf324ce44/resourceGroups/dtolab-sitea-rg/providers/Microsoft.Security/locations/eastus/jitNetworkAccessPolicies/default'
                    Start-AzJitNetworkAccessPolicy -ResourceId "/subscriptions/$SubscriptionID/resourceGroups/$($VM.ResourceGroup)/providers/Microsoft.Security/locations/$($VM.Location)/jitNetworkAccessPolicies/default" `
                                                    -VirtualMachine $JitPolicyArr -ErrorAction Stop | Out-Null
                    If($ShowStatus){Write-Host 'Success' -ForegroundColor Green}
                }
                Catch{
                    If($VerbosePreference -eq 'Continue'){
                        Write-Host ("Failed to request JIT access on VM [{0}]. Error: {1}{2}" -f $VM.VMName,$_.Exception.ItemName,$_.Exception.Message) -ForegroundColor Red
                    }Else{
                        If($ShowStatus){Write-Host ("Failed to request JIT access on VM [{0}]." -f $VM.VMName) -ForegroundColor red}
                    }
                }
            }
            Else{
                If($ShowStatus){Write-Host 'Exists' -ForegroundColor Green}
            }
        }#end loop
    }
    End{
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            Foreach($VM in $VMName)
            {
                $global:MyAzVMs = Get-MyAzureVM -VMName $VM -NetworkDetails -NoStatus
            }
        }
        Else{
            $global:MyAzVMs = Get-MyAzureVM -NetworkDetails
        }

        If($ShowStatus){
            $global:MyAzVMs | Select-Object VMName,HostName,LocalIP,PublicIP,PublicDNS,State,JITAccess | Format-Table
        }
        Stop-Transcript | Out-Null

    }
}
Function Enable-MyAzureJitPolicy{
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [Alias("VM")]
        [string[]]$VMName,

        [Parameter(Mandatory = $False,
            ValueFromPipelineByPropertyName = $true)]
        [Alias("ResourceGroup")]
        [string]$ResourceGroupName = $global:MyResourceGroup
    )
    Begin
    {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #attempt connection to Azure. If connection is not successfull clear all settings and try again
        Try{
            $Null = Connect-MyAzureEnvironment -ErrorAction Stop
        }
        Catch{
            Connect-MyAzureEnvironment -ClearAll
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
            $global:MyResourceGroup = $ResourceGroupName
        }

        #if resource group is blank, throw an error
        If($null -eq $ResourceGroupName){
            Throw "Resource group name has not been specified"
        }
    }
    Process
    {
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            Foreach($VM in $VMName)
            {
                $VMs += Get-MyAzureVM -VMname $VM -NetworkDetails -NoStatus
            }
        }
        Else{
            $VMs = Get-MyAzureVM -NetworkDetails -NoStatus
        }

        Foreach($VM in $VMName){
            #Assign a variable that holds the just-in-time VM access rules for a VM:
            $JitPolicy = (@{
                id="/subscriptions/$global:MySubscriptionID/resourceGroups/$ResourceGroupName/providers/Microsoft.Compute/virtualMachines/$VM";
                ports=(@{
                     number=22;
                     protocol="*";
                     allowedSourceAddressPrefix=@("*");
                     maxRequestAccessDuration="PT3H"},
                     @{
                     number=3389;
                     protocol="*";
                     allowedSourceAddressPrefix=@("*");
                     maxRequestAccessDuration="PT3H"})})
            #Insert the VM just-in-time VM access rules into an array
            $JitPolicyArr=@($JitPolicy)
            #Configure the just-in-time VM access rules on the selected VM:
            Set-AzJitNetworkAccessPolicy -Kind "Basic" -Location $VM.Location -Name "default" -VirtualMachine $JitPolicyArr
            #$jitString = ('id=' + $JitPolicyArr.id + ';ports={maxRequestAccessDuration=' + $JitPolicyArr.ports.maxRequestAccessDuration[0] + '}')
            #If($VerbosePreference -eq 'Continue'){Write-Host "Command: Set-AzJitNetworkAccessPolicy -Kind "Basic" -Location $($VM.Location) -VirtualMachine {$jitString}" -ForegroundColor Yellow}
        }
    }
    End{Stop-Transcript | Out-Null}
}

Function Start-MyLabEnvironment{
    ## A Simple one line function to do it all
    Write-host ("Preparing Lab environment...") -ForegroundColor Yellow

    Set-MyAzureEnvironment -MyEnv $global:MyLabTenant

    Start-MyAzureEnvironment -OrderTag $global:MyLabTag

    $LocalGateway = Get-AzLocalNetworkGateway
    If($LocalGateway.GatewayIpAddress -ne $global:MyPublicIP)
    {
        Write-host ("Azure Gateway IP does not match your public IP, updating to: ") -NoNewline
        Write-host ("{0}..." -f $global:MyPublicIP) -ForegroundColor Cyan -NoNewline
        Try{
            New-AzLocalNetworkGateway -Name $LocalGateway.Name -Location $LocalGateway.Location -AddressPrefix $LocalGateway.LocalNetworkAddressSpace.AddressPrefixes `
                -GatewayIpAddress $global:MyPublicIP -ResourceGroupName $LocalGateway.ResourceGroupName -Force
            Write-host ("Success") -ForegroundColor Green
        }Catch{
            Write-host ("Failed: {0}" -f $_.exception.message) -ForegroundColor Red
        }
    }
    Else{
        Write-host ("Azure Gateway IP is accurate and is set to: ") -NoNewline
        Write-host ("{0}..." -f $global:MyPublicIP) -ForegroundColor Green
        Write-host ("Checking gateway connection status...")
        $Gateways = Get-AzVirtualNetworkGatewayConnection -ResourceGroupName $LocalGateway.ResourceGroupName
        #TEST $Gateway = $Gateways[0]
        Foreach($Gateway in $Gateways){
            $GatewayConnection = Get-AzVirtualNetworkGatewayConnection -name $Gateway.Name -ResourceGroupName $LocalGateway.ResourceGroupName
            Write-host ("Gateway [{0}] status is: " -f $Gateway.Name) -NoNewline
            If($GatewayConnection.connectionStatus -eq 'Connected')
            {
                Write-host ("{0}" -f $GatewayConnection.connectionStatus) -ForegroundColor Green
                $VPNConnected = $true
            }
            Else{
                Write-host ("{0}" -f $GatewayConnection.connectionStatus) -ForegroundColor Red
                $VPNConnected = $false
            }
        }
    }

    If((Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online).State -eq 'Enabled'){
        Write-host ("Please wait while starting all Hyper-V VM's by tag order...")
        Start-MyHyperVM -OrderTag $global:MyLabTag
    }

    If(Test-Connection $global:MyLabRouterIP -Count 1){
        Write-host ("Local Lab router is running") -NoNewline -ForegroundColor Gray
        If($VPNConnected -eq $false){Write-host "...SSH to $global:MyLabRouterIP and run 'restart vpn'"  -ForegroundColor Yellow}
    }Else{
        Write-host ("Local Lab router is not running and connected to VPN") -ForegroundColor Red

    }

    Write-Host ""
    Write-Host "Connect to Azure VM:" -ForegroundColor Green
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host "ConnectTo-MyAzureVM" -ForegroundColor Cyan -NoNewline
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-VMName"-ForegroundColor White -NoNewline
    Write-Host " <tab to ip or dns>"-ForegroundColor gray -NoNewline
    Write-Host "]"-ForegroundColor darkgray
    Write-Host ("  eg. ConnectTo-MyAzureVM -VMname azurevm1.eastus.cloudapp.azure.com") -ForegroundColor DarkGray
    Write-Host "==================================================================" -ForegroundColor Green
}



Function Start-MyAzureEnvironment{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true)]
        [Alias("ResourceGroup")]
        [string]$ResourceGroupName = $global:MyResourceGroup,
        [string]$OrderTag,
        [boolean]$OutVoice = $VoiceWelcomeMessage
    )
    Begin
    {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #attempt connection to Azure. If connection is not successfull clear all settings and try again
        Try{
            $Null = Connect-MyAzureEnvironment -ErrorAction Stop
        }
        Catch{
            Connect-MyAzureEnvironment -ClearAll
        }

        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
            $global:MyResourceGroup = $ResourceGroupName
        }

        <#
        #overwrite global resource group is parameter is called
        if ($PSBoundParameters.ContainsKey('ResourceGroupName')) {
            $global:MyResourceGroup = $ResourceGroupName
        }Else{
            $ResourceGroupName = $global:MyResourceGroup
        }
        #>

        #if resource group is blank, throw an error
        If($null -eq $ResourceGroupName){
            Throw "Resource group name has not been specified"
        }

        #TEST $OrderTag = 'StartupOrder'
        $VMs = @()
    }
    Process
    {
        If($PSBoundParameters.ContainsKey('OrderTag') )
        {
            Write-host ("Please wait while collecting all Azure VM's by tag order...") -NoNewline
            $VMs = Start-MyAzureVM -OrderTag $OrderTag -NoStatus
        }
        Else{
            $message = ("Please wait while collecting all Azure VM's")
            Write-Host ($message + '...') -ForegroundColor Gray -NoNewline
            If($OutVoice){Out-MyVoice $message}
            $VMs = Get-MyAzureVM -NetworkDetails -NoStatus

            #Write output cleaniung based on number
            If($VMs.count -eq 1){Write-host ("[{0}] vm found" -f $VMs.count) -ForegroundColor Green}
            Else{Write-host ("[{0}] VM's found" -f $VMs.count) -ForegroundColor Green}

            #TEST $VM = $VMs[0]
            foreach ($VM in $VMs)
            {
                Write-host ("Ensuring VM: ") -ForegroundColor Gray -NoNewline
                Write-host ("{0}" -f $VM.VMName) -NoNewline
                Write-host (" is ready to be used..." ) -ForegroundColor Gray
                Try{
                    Write-host ("  Checking power state...") -ForegroundColor Gray -NoNewline

                    If($VM.State -ne 'Running'){
                        Write-host ("starting") -ForegroundColor Green

                        #If($VerbosePreference -eq 'Continue'){Write-Host "Command: Start-MyAzureVM -VMName $($VM.VMName) -NoStatus" -ForegroundColor Yellow}
                        $vmstate = Start-MyAzureVM -VMName $VM.VMName -NoStatus
                    }Else{
                        Write-host ("running") -ForegroundColor Green
                    }
                }
                Catch{
                    If($VerbosePreference -eq 'Continue'){
                        Write-host ("failed to start: {0}" -f $_.Exception.ItemName) -ForegroundColor red
                    }Else{
                        Write-host ("failed to start") -ForegroundColor red
                    }
                }

            }#end loop
        }
    }
    End{
        foreach ($VM in $Global:MyAzVMs)
        {
            If( ($VM.JITAccess -ne 'disabled' -or $VM.JITAccess -eq 'False') -and $VM.PublicIP ){
                Try{
                    Write-host ("  Checking Just-In-Time policy on VM: ") -ForegroundColor Gray -NoNewline
                    Write-host ("{0}" -f $VM.VMName) -NoNewline -ForegroundColor Cyan
                    Write-host (" for public IP: ") -ForegroundColor Gray -NoNewline
                    Write-host ("{0}" -f $global:MyPublicIP) -NoNewline -ForegroundColor Cyan
                    Write-host ("...") -ForegroundColor Gray -NoNewline
                    #If($VerbosePreference -eq 'Continue'){Write-Host "Command: Set-MyAzureJitPolicy -VMName $($VM.VMName) -NoStatus" -ForegroundColor Yellow}
                    $jitaccess = Set-MyAzureJitPolicy -VMName $VM.VMName -NoStatus

                    If($jitaccess.JITAccess -eq $True){
                        Write-host 'allowed' -ForegroundColor Green
                    }Else{
                        Write-host 'enabling' -ForegroundColor Yellow
                    }
                }
                Catch{
                    If($VerbosePreference -eq 'Continue'){
                        Write-host ("failed to enable JIT: {0}" -f $_.Exception.ItemName) -ForegroundColor red
                    }Else{
                        Write-host ("failed enable JIT") -ForegroundColor red
                    }
                }
            }
        }#end loop

        $global:MyAzVMs = Get-MyAzureVM -NetworkDetails
        $global:MyAzVMs | Select-Object VMName,HostName,LocalIP,PublicIP,PublicDNS,State,JITAccess | Format-Table

        Stop-Transcript | Out-Null
    }
}


Function Convert-XMLtoPSObject {
    Param (
        $XML
    )
    $Return = New-Object -TypeName PSCustomObject
    $xml |Get-Member -MemberType Property |Where-Object {$_.MemberType -EQ "Property"} |ForEach-Object {
        IF ($_.Definition -Match "^\bstring\b.*$") {
            $Return | Add-Member -MemberType NoteProperty -Name $($_.Name) -Value $($XML.($_.Name))
        } ElseIf ($_.Definition -Match "^\System.Xml.XmlElement\b.*$") {
            $Return | Add-Member -MemberType NoteProperty -Name $($_.Name) -Value $(Convert-XMLtoPSObject -XML $($XML.($_.Name)))
        } Else {
            Write-Host " Unrecognized Type: $($_.Name)='$($_.Definition)'"
        }
    }
    $Return
}

Function ConnectTo-MyAzureVM{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [ArgumentCompleter( {
            param ( $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )

            $RemoteVMs = @()
            $RemoteVMs += $Global:MyAzVMs | Where-Object {$Null -ne $_.PublicDNS} | Select-Object -ExpandProperty PublicDNS
            $RemoteVMs += $Global:MyAzVMs | Where-Object {$Null -ne $_.PublicIP} | Select-Object -ExpandProperty PublicIP

            $RemoteVMs | Where-Object {
                $_ -like "$wordToComplete*"
            }

        } )]
        [Alias("VM")]
        [string]$VMname,
        [string]$ResourceGroupName = $global:MyResourceGroup,
        [switch]$RequestJIT,
        [Alias("domainUser")]
        [string]$Username,
        [System.Management.Automation.Credential()]$Credential

    )

    If($PSBoundParameters.ContainsKey('RequestJIT') ){
       Write-host ("Attempting JIT Approval request for {0}..." -f $VMName)
       Set-MyAzureJitPolicy -VMName $VMName
    }

    # initiate the RDP connection
    # connection will automatically use cached credentials
    # if there are no cached credentials, you will have to log on
    # manually, so on first use, make sure you use -Credential to submit
    Write-host ("Attempting RDP connection to {0}.." -f $VMName)
    Try{
        if ($PSBoundParameters.ContainsKey('Credential') )
        {
            # extract username and password from credential
            $User = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password

            # save information using cmdkey.exe
            Start-Process cmdkey -ArgumentList "/generic:$VMName /user:$User /pass:$Password" -Wait -ErrorAction Stop
        }
        else{
            Start-Process cmdkey -ArgumentList "/generic:$VMName /user:$Username" -Wait -ErrorAction Stop
        }

        Start-Process mstsc -ArgumentList "/v:$VMName /f" -NoNewWindow -PassThru -ErrorAction Stop
        Start-Sleep 30
    }
    Catch{
        Write-host ("Failed to connect to {0}: {1}" -f $VMName, $_.Exception.Message)
    }
    Finally{
        Start-Process cmdkey -ArgumentList "/delete:$VMName" -Wait
    }
}

Function Get-MyRemoteDesktopData{
    $Path = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftRemoteDesktopPreview_8wekyb3d8bbwe\LocalState\RemoteDesktopData\LocalWorkspace\connections"
    If(Test-Path $Path -ErrorAction SilentlyContinue){
        $RDPConnFiles = Get-Childitem $Path -Filter *.model
        $Content = @()
        Foreach($Model in $RDPConnFiles){
            $Data = (Get-content $model.Fullname) -replace 'a:','' -replace 'i:type="ConnectionArgsModel"', ''
            [xml]$XMLData = $Data
            $Content += (Convert-XMLtoPSObject $XMLData).SerializableModel
        }
    }
    return $Content
}


Function Get-MyAzureUserName{
    Param (
        [switch]$firstname
	)
    If(Test-Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -ErrorAction SilentlyContinue){
        $IDDisplayName = Get-ItemPropertyValue "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name ADUserDisplayName
        If($null -ne $IDDisplayName){
            $Name = $IDDisplayName
        }Else{
            $Name = $Env:UserName
        }
    }
    ElseIf(Test-Path "$env:LOCALAPPDATA\.IdentityService\V2AccountStore.json" -ErrorAction SilentlyContinue){
        $IDStore = Get-content "$env:LOCALAPPDATA\.IdentityService\V2AccountStore.json"
        $IDPayload = ($IDStore | ConvertFrom-Json).Properties.IdTokenPayload
        If($null -ne $IDPayload){
            $Name = ($IDPayload | ConvertFrom-Json).name | Select-Object -Last 1
        }
    }
    Else{
        $Name = $Env:UserName
    }

    If($firstname){
        $Name = $Name.Split(' ')[0]
    }
    Return $Name
}

Function Get-MyRandomAlphanumericString {
	[CmdletBinding()]
	Param (
        [int] $length = 8
	)

	Begin{
	}
	Process{
        return ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) | Get-Random -Count $length  | ForEach-Object {([char]$_).ToString().ToUpper()}) )
	}
}


Function Get-MyRandomSerialNumber{
    "$(Get-MyRandomAlphanumericString -length 3)$(Get-random -Minimum 1000000 -Maximum 9999999)$(Get-MyRandomAlphanumericString -length 2)"

}

Function Start-MyHyperVM{
# Microsoft - Compute Resources
#-----------------------------
   [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [Alias("VM")]
        [string[]]$VMName,
        [string]$OrderTag,
        [switch]$NoStatus
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null

        #default to showing status (like verbose) unless noStatus is in param
        if ($PSBoundParameters.ContainsKey('NoStatus')) {
            $ShowStatus=$false
        }
        else{
            $ShowStatus=$true
        }

        $VMs = @()
    }
    Process{

        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
             #TEST $VM = $VMname
            Foreach($VM in $VMName)
            {
                $startedmessage = ("{0} is started already!" -f $VM)
                $VMs += Get-MyHyperVM -VMName $VM -NetworkDetails -NoStatus
            }
        }
        Else{
            $startedmessage = "All VM's are started already!"
            $VMs = Get-MyHyperVM -NetworkDetails -NoStatus
        }

        If($ShowStatus){Write-Host 'Collecting VM running status...' -ForegroundColor DarkGray -NoNewline}

        If($VMs.count -eq 0)
        {
            If($ShowStatus){Write-Host $startedmessage -ForegroundColor Green}
        }
        Else{
            If( $PSBoundParameters.ContainsKey('OrderTag') ){
                # Get the StartupOrder tag, if missing set to be run last (10)
                $taggedVMs = @{}

                #TEST $OrderTag = 'StartupOrder'
                #TEST $VMs | Select-Object name,notes
                #TEST $vm = $VMs[1]
                #first grab vms with tags
                ForEach ($vm in $VMs) {
                    #$startupValue = $null

                    if ($vm.Tags[$OrderTag] -eq 0)
                    {
                        Continue
                        #do not add vm to list
                        #$startupValue = $vm.Tags['StartupOrder']
                    }
                    Else
                    {
                        If($VerbosePreference){write-host ('Found {0} tag for {1}; value is: {2}' -f $OrderTag,$vm.Name,$vm.Tags[$OrderTag])}
                        $taggedVMs.Add($vm.Name,$vm.Tags[$OrderTag])
                    }

                    #$taggedVMs.Add($vm.name,$startupValue)
                }


                #Count must be more that 0 to continue
                If($taggedVMs.count -gt 0)
                {
                    #get max value of startup count and make that the start if the next loop
                    $StartupEndCount = $taggedVMs.values | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum


                    #increment number to tag that are null
                    foreach($key in $taggedVMs.Keys.Clone())
                    {
                        If ($null -eq $taggedVMs[$key])
                        {
                            $StartupEndCount = $StartupEndCount + 1
                            If($VerbosePreference){write-host ('Adding tag: {0}={1}' -f $key,$StartupEndCount)}
                            $taggedVMs[$key] += $StartupEndCount
                        }
                    }

                    Write-Host ("[{0}] VM's tagged to start" -f $taggedVMs.count) -ForegroundColor Green

                    # Start in order from lowest vm
                    $tagOrderNum = $taggedVMs.values | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
                    $i = 0
                    #TEST ITERATION: $tagOrderNum = 2
                    Do{
                        #increment iteration count
                        $i++
                        #Always null VM to start
                        $tobeStarted = $null
                        # Get the VM tag that matched current iteration
                        $tobeStarted = $taggedVMs.GetEnumerator().Where({$_.Value -eq $tagOrderNum}) | Select-Object -ExpandProperty Key
                        If($tobeStarted)
                        {
                            Write-Host ("  Attempting to start VM in order: {0}" -f $tobeStarted)
                            Start-VM -Name $tobeStarted -asJob | Out-Null

                        }
                        Else{
                            Write-Host ("  No VM found with {0}: {1}" -f $OrderTag,$tagOrderNum)
                        }
                        #increment tag order
                        $tagOrderNum++
                    }
                    Until ($i -gt $taggedVMs.count)

                }
                Else{
                    Write-Host ("No running VM's were tagged to start") -ForegroundColor Yellow
                }
            }
            Else{
                #start all deallocated VM's
                foreach ($VM in $VMs)
                {
                    If($ShowStatus){write-host ("Attempting to start VM: {0}" -f $VM.Name)}
                    Start-VM -Name $VM.Name -asJob | Out-Null
                }
            }
        }
    }
    End{
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            Foreach($VM in $VMName)
            {
                $global:MyVMs = Get-MyHyperVM -Name $VM -NetworkDetails -NoStatus
            }
        }
        Else{
            $global:MyVMs = Get-MyHyperVM -NetworkDetails | Where-Object {$_.State -eq 'Running'}
        }
        If($ShowStatus){
            $global:MyVMs | Select-Object Name,LocalIP,State | Format-Table
        }
        Stop-Transcript | Out-Null
    }
}


Function Get-MyHyperVM{
    [CmdletBinding(DefaultParameterSetName = 'ListParameterSet',
        HelpUri = 'https://go.microsoft.com/fwlink/?LinkID=398573',
        SupportsShouldProcess = $true)]
    Param(
        [Parameter(Mandatory = $false,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName = $true,
            Position = 0,
            ParameterSetName = 'VMParameterSet')]
        [ArgumentCompleter( {
            param ( $commandName,
                    $parameterName,
                    $wordToComplete,
                    $commandAst,
                    $fakeBoundParameters )


            Get-VM -ErrorAction Silentlycontinue | ForEach-Object {$_.Name} | Where-Object {
                $_ -like "$wordToComplete*"
            }

        } )]
        [Alias("VM")]
        [string]$VMname,
        [switch]$NetworkDetails,
        [switch]$NoStatus
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        #build log name
        [string]$FileName = 'Profile_' + ${CmdletName} + '_' + (get-date -Format MM-dd-yyyy) + '.log'
        Start-Transcript -Path $env:TEMP\$FileName -Force -Append | Out-Null


        #default to showing status (like verbose) unless noStatus is in param
        if ($PSBoundParameters.ContainsKey('NoStatus')) {
            $ShowStatus=$false
        }
        else{
            $ShowStatus=$true
        }

        #show the appropiate message
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            If($ShowStatus){Write-Host ("Collecting Hyper-V VM with names [{0}]..." -f ($VMName -join ",")) -ForegroundColor DarkGray -NoNewline}
        }Else{
            If($ShowStatus){Write-Host "Collecting all Hyper-V VM's..." -ForegroundColor DarkGray -NoNewline}
        }

        $VMs = @()
        $report = @()

        #pull all Virtual Machine's; this is easier and faster than build VM's list one at a time.
        $AllVMs = Get-VM -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        $vmwp = Get-CimInstance Win32_Process -Filter "Name like '%vmwp%'"
        #$vmwp | ForEach-Object {$_.CommandLine.split(" ")[1]}
    }
    Process{
        If($PSCmdlet.ParameterSetName -eq "VMParameterSet"){
            #filter out only listed VM's
            #TEST $VM = $VMname
            Foreach($VM in $VMName)
            {
                $VMs += $AllVMs | Where-Object {$_.Name -eq $VM}
            }
            If($ShowStatus){Write-Host 'completed' -ForegroundColor Green}
        }
        Else{
            $VMs = $AllVMs
            If($ShowStatus){Write-Host 'completed' -ForegroundColor Green}
        }

        if ($PSBoundParameters.ContainsKey('NetworkDetails'))
        {
            #list VM and their network info
            #TEST $vm = $VMs[3]
            #TEST $vm = $VMs[6]
            #TEST $vm = $VMs[11]
            foreach($vm in $VMs)
            {
                If($ShowStatus){Write-Host '.' -ForegroundColor DarkGray -NoNewline}

                $vmwp | ForEach-Object {
                    #Try{
                        $Process = $_ | Where-Object {$_.CommandLine.split(" ")[1].Trim() -eq $VM.id.guid.ToUpper()}
                    #}Catch{}
                }

                #grab only IPv4address as a comma list
                $Octet = '(?:0?0?[0-9]|0?[1-9][0-9]|1[0-9]{2}|2[0-5][0-5]|2[0-4][0-9])'
                [regex]$IPv4Regex = "^(?:$Octet\.){3}$Octet$"
                $IPv4Addresses = [regex]::Matches($VM.NetworkAdapters.IpAddresses,$IPv4Regex).Value -join ','

                #grab tags in notes
                $VMTags = @{}
                $VM.Notes -split '\n' | ForEach-Object { $s = $_ -split ':'; $VMTags += @{$s[0].Trim() =  $s[1].Trim()}}

                $info = "" | Select-Object ProcessID,Id,Name,LocalIP,State,Tags
                $info.ProcessID = $Process.ProcessID
                $info.Id = $VM.id.guid.ToUpper()
                $info.Name = $VM.Name
                $info.LocalIP = $IPv4Addresses
                $info.State = $VM.state
                $info.Tags = $VMTags

                $report+=$info
            }

            return $report
        }
        Else{
            return $VMs
        }
    }
    End{Stop-Transcript | Out-Null}

}

Function Stop-MyHyperVM{
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string]$Id
    )
    Begin{}
    Process{
        Get-Process -Id $Id | Stop-Process -Force -Verbose
    }
}

Function Restart-MyHyperV{
    #$guid = Get-VM dtolab-ap1 | Select-Object -ExpandProperty vmid
    Get-Process vmms | Stop-Process -Force
    Start-Sleep 10
    Get-Service vmms | Start-Service
}

Function Set-MyWindowPosition {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $Process,

        [parameter(Mandatory=$false)]
        [ValidateSet('Hide','Minimize','Maximize','Restore')]
        [string]$Position,

        [switch]$Show
    )
    Begin{
        Try{
            $WindowCode = '
                [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
                [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
            '
            $AsyncWindow = Add-Type -MemberDefinition $WindowCode -Name Win32ShowWindowAsync -namespace Win32Functions -PassThru
        }
        Catch{

        }
        Finally{
            $hwnds = @($Process)
        }
    }
    Process{
        Foreach($hwnd in $hwnds)
        {
            switch($Position)
            {

             # hide the window (remove from the taskbar)
            'Hide'      {$null = $AsyncWindow::ShowWindowAsync($hwnd.MainWindowHandle, 0)}

            'Maximize'  {$null = $AsyncWindow::ShowWindowAsync($hwnd.MainWindowHandle, 3)}

             #open window
             #{$null = $AsyncWindow::ShowWindowAsync($hwnd0.MainWindowHandle, 4)}

            'Minimize'  {$null = $AsyncWindow::ShowWindowAsync($hwnd.MainWindowHandle, 6)}
                    # restore the window to its original state
            'Retore'    {$null = $AsyncWindow::ShowWindowAsync($hwnd.MainWindowHandle, 9)}
            }

            # being in front
            If($Show){
                $null = $AsyncWindow::SetForegroundWindow($hwnd.MainWindowHandle)
            }
        }
    }End{

    }
}

Function Show-MyCommands
{
    Write-Host ""
    Write-Host "Custom functions available:" -ForegroundColor Green
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host "Connect-MyAzureEnvironment" -ForegroundColor Cyan -NoNewline
    Write-Host " -TenantId" -ForegroundColor White -NoNewline
    Write-Host " <tenant guid>" -ForegroundColor Gray
    Write-Host ("  eg. Connect-MyAzureEnvironment -TenantId {0} -SubscriptionId {1}" -f (New-Guid).Guid,(New-Guid).Guid) -ForegroundColor DarkGray
    Write-Host "Get-MyAzureVM" -ForegroundColor Cyan -NoNewline
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-VMName"-ForegroundColor White -NoNewline
    Write-Host " <vm(s)>"-ForegroundColor gray -NoNewline
    Write-Host "]"-ForegroundColor darkgray -NoNewline
    Write-Host " -NetworkDetails" -ForegroundColor White
    Write-Host "  eg. Get-MyAzureVM -VMName DC01,CA01,PR01 -NetworkDetails" -ForegroundColor DarkGray
    Write-Host "Start-MyAzureVM" -ForegroundColor Cyan -NoNewline
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-VMName"-ForegroundColor White -NoNewline
    Write-Host " <vm(s)>"-ForegroundColor gray -NoNewline
    Write-Host "]"-ForegroundColor darkgray
    Write-Host "  eg. Start-MyAzureVM -VMName DC01,CA01,PR01" -ForegroundColor DarkGray
    Write-Host "Set-MyJitAccess" -ForegroundColor Cyan -NoNewline
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-VMName"-ForegroundColor White -NoNewline
    Write-Host " <vm(s)>"-ForegroundColor gray -NoNewline
    Write-Host "]"-ForegroundColor darkgray
    Write-Host "  eg. Set-MyAzureJitPolicy -VMName DC01,CA01,PR01" -ForegroundColor DarkGray
    Write-Host "Set-MyAzureEnvironment" -ForegroundColor Cyan -NoNewline
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-MyEnv"-ForegroundColor White -NoNewline
    Write-Host " <tab to env>"-ForegroundColor gray -NoNewline
    Write-Host "]"-ForegroundColor darkgray
    Write-Host "  eg. Set-MyAzureEnvironment -MyEnv SiteA" -ForegroundColor DarkGray
    Write-Host "Start-MyAzureEnvironment" -ForegroundColor Cyan -NoNewline
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-OrderTag"-ForegroundColor White -NoNewline
    Write-Host " 'StartupOrder'"-ForegroundColor gray -NoNewline
    Write-Host "]"-ForegroundColor darkgray
    Write-Host " ["-ForegroundColor darkgray -NoNewline
    Write-Host "-IncludeLocalVM"-ForegroundColor White -NoNewline
    Write-Host "]"-ForegroundColor darkgray
    Write-Host "  eg. Start-MyAzureEnvironment -OrderTag 'StartupOrder'" -ForegroundColor DarkGray
    Write-Host "Start-MyLabEnvironment" -ForegroundColor Cyan
    Write-Host "  NOTE. This uses global variables in script" -ForegroundColor DarkGray
    Write-Host "==================================================================" -ForegroundColor Green
    Write-Host 'Your Public IP is: ' -ForegroundColor Gray -NoNewline
    Write-Host $global:MyPublicIP -ForegroundColor Cyan
    Write-Host ""
}

#=======================================================
# MAIN
#=======================================================
#exit process of profile script if using vscode
if (Test-MyVSCode) { exit }

$Hour = (Get-Date).Hour
$UserName = Get-MyAzureUserName -firstname
If($VoiceWelcomeMessage){
    If ($Hour -lt 10) {("Good Morning, {0}" -f $UserName) | Out-MyVoice -PassThru}
    ElseIf ($Hour -gt 16) {("Good Evening, {0}" -f $UserName) | Out-MyVoice -PassThru}
    Else {("Good Afternoon, {0}" -f $UserName)| Out-MyVoice -PassThru}
}
Else{
    If ($Hour -lt 10) {Write-Host ("Good Morning, {0}" -f $UserName)}
    ElseIf ($Hour -gt 16) {Write-Host ("Good Evening, {0}" -f $UserName)}
    Else {Write-Host ("Good Afternoon, {0}" -f $UserName)}
}
Write-Host 'Your running Powershell version: ' -ForegroundColor Gray -NoNewline
Write-Host  $PsVersionTable.PSVersion -ForegroundColor Cyan
'-----------------------------------------------'
If($VoiceWelcomeMessage){"Please wait while I check for installed modules" | Out-MyVoice}
Install-MyLatestModule -Name $Checkmodules -Frequency Daily

Show-MyCommands

#create a vsc alias like ise
If(Test-MyVSCodeInstall){
    Set-Alias vsc -Value code
}Else{
    Write-host "Visual Studio Code was not found; install at https://code.visualstudio.com/ or run command Start-MyVSCodeInstall" -BackgroundColor Red -ForegroundColor White
}

#Open-MyFile -filename "$env:USERPROFILE\Downloads\CA01.rdp" -method run