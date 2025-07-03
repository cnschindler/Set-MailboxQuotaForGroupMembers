#############################
#
# Set-MailboxQuotasForGroupMembers.ps1
# Resets Users to default Quota if they are
# not Member of a Custom Quota Group
#
# Sets Quota of Groupmembers to predefined Values
#

#Requires -Version 3.0
#Requires -Module ActiveDirectory

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [System.IO.FileInfo]
    $ConfigFile = (Join-Path -Path $PSScriptRoot -ChildPath "Set-MailboxQuotaForGroupMembers_Config.json")
)

$Config = Get-Content -Path $ConfigFile | ConvertFrom-Json

$Domaincontroller=$Config.Domaincontroller
$domain = $Config.domain
#$LogFileAge = $Config.LogFileAge
$Filter = $config.Filter -join ' -and '
$Quotas = $Config.Quotas

[string]$LogfileFullPath = Join-Path -Path $PSScriptRoot (Join-Path $MyInvocation.MyCommand.Name ($MyInvocation.MyCommand.Name + "_" + $($ContactSourceMailbox.Split("@")[0]) + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now))
$script:NoLogging = $Config.NoLogging
function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured variables 'LogfileFullPath' to identify the logfile and 'NoLogging' to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$LogPrefix,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )

    # Prefix the string to write with the current Date and Time, add error message if present...
    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : ERROR {1}: {2} Error: {3}" -f [DateTime]::Now, $LogPrefix, $Message, $ErrorInfo.Exception.Message
    }

    else
    {
        $logLine = "{0:d.M.y H:mm:ss} : INFO {1}: {2}" -f [DateTime]::Now, $LogPrefix, $Message
    }

    if (-not $NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $LogfileFullPath -PathType Leaf))
        {
            New-Item -ItemType File -Path $LogfileFullPath -Force -Confirm:$false -WhatIf:$false | Out-Null
            Add-Content -Value "Logging started." -Path $LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    else
    {
        Write-Host $logLine
    }
}
Function LoadADModule
{
    $ModuleName = "ActiveDirectory"
    $IsModuleInstalled = (Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1)
    
    if ($IsModuleInstalled.Name -eq "$($ModuleName)")
    {   
        try
        {
            Import-Module -Name $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
            Write-LogFile -LogPrefix "LoadADModule" -Message "ActiveDirectory Module successfully loaded."
        }
        
        catch
        {
            $Textbox_Messages.Text = "ActiveDirectory Module could not be loaded. Error: $($Error.Exception.InnerException)"
            Write-LogFile -LogPrefix "LoadADModule" -Message "ActiveDirectory Module could not be loaded." -ErrorInfo $_}
    }

    else
    {
        Write-LogFile -LogPrefix "LoadADModule" -Message "ActiveDirectory Module not installed. Please install first!"
    }
} 
function ConnectExchange
{
    # Check if a connection to an exchange server exists and connect if necessary...
    if (-NOT (Get-PSSession | Where-Object ConfigurationName -EQ "Microsoft.Exchange"))
    {
        $LogPrefix = "ConnectExchange"

        # Test if Exchange Management Shell Module is installed - if not, exit the script
        $EMSModuleFile = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction SilentlyContinue).MsiInstallPath + "bin\RemoteExchange.ps1"
        
        # If the EMS Module wasn't found
        if (-Not (Test-Path $EMSModuleFile))
        {
            # Write Error end exit the script
            $ErrorMessage = "Exchange Management Shell Module not found on this computer. Please run this script on a computer with Exchange Management Tools installed!"
            Write-LogFile -LogPrefix $LogPrefix -Message $ErrorMessage
            Exit
        }

        else
        {
            # Load Exchange Management Shell
            try
            {
                # Dot source the EMS Script
                . $($EMSModuleFile) -ErrorAction Stop | Out-Null
                Write-LogFile -LogPrefix $LogPrefix -Message "Successfully loaded Exchange Management Shell Module."
            }

            catch
            {
                Write-LogFile -LogPrefix $LogPrefix -Message "Unable to load Exchange Management Shell Module." -ErrorInfo $_
                Exit
            }

            # Connect to Exchange Server
            try
            {
                Connect-ExchangeServer -auto -ClientApplication:ManagementShell -ErrorAction Stop | Out-Null
                Write-LogFile -LogPrefix $LogPrefix -Message "Successfully connected to Exchange Server."
            }

            catch
            {
                Write-LogFile -LogPrefix $LogPrefix -Message "Unable to connect to Exchange Server." -ErrorInfo $_
                Exit
            }
        }
    }
}

#############################
# Custom Object Decleration
#############################

$GroupLimits = @()
$newline = New-Object –TypeName PSCustomObject
Add-Member -InputObject $newline -MemberType NoteProperty -Name Group -Value ""
Add-Member -InputObject $newline -MemberType NoteProperty -Name Warning -Value ""
Add-Member -InputObject $newline -MemberType NoteProperty -Name Send -Value ""
Add-Member -InputObject $newline -MemberType NoteProperty -Name Receive -Value ""

foreach($line in $mygroupandlimits)
{
    $newline.Group = $line[0]
    $newline.Warning = $line[1]
    $newline.Send = $line[2]
    $newline.Receive = $line[3]

    $GroupLimits += ($newline | select *)
}

######################
# Get Quotagroupmembers
######################

$MBXGroup = @()

foreach($group in $GroupLimits)
{
    foreach($Sam in ((Get-ADGroupMember $group.Group -Server $Domaincontroller).SamAccountName)){
    $MBXGroup += get-mailbox $domain$sam -DomainController $Domaincontroller #-ErrorAction silentlycontinue
    }
}



$MBXQuota = Get-Mailbox -resultsize unlimited -DomainController $Domaincontroller -filter $Filter

$MBXComp = (Compare-Object -ReferenceObject $MBXGroup.samaccountname -DifferenceObject $MBXQuota.samaccountname | ? {$_.SideIndicator -eq "=>"})


######################
# User is not Member of a Custom Quota Group -> Reset to Database Default !
######################
foreach($mbx in $MBXComp)
{
#    Set-Mailbox -DomainController $Domaincontroller -IssueWarningQuota unlimited -ProhibitSendQuota unlimited -ProhibitSendReceiveQuota unlimited -UseDatabaseQuotaDefaults $true -Identity $domain$($mbx.InputObject)
#    Set-Mailbox -DomainController $Domaincontroller -UseDatabaseQuotaDefaults $true -Identity $domain$($mbx.InputObject)
    $message = "Set DatabaseDefaultLimits to User " + $mbx.InputObject + "."
    Write-Log -Message $message -Path $logfile
}




## Set Quota
## 
######################
# Set Custom Quota for every Member of the Groups
######################

foreach($group in $GroupLimits)
{
    
    $changefile = ".\" + $group.Group + ".csv"

	if(Test-Path $changefile)
	{
		$changelist = Import-Csv -Delimiter "|" -Encoding Unicode -Path $changefile
	}else
    {
        $changelist = New-Object –TypeName PSCustomObject
        Add-Member -InputObject $changelist -MemberType NoteProperty -Name SamAccountName -Value "nomembers_in_AD_Group_for_custom_quota"
    }
    
    $MBX2Change = @()

    foreach($sam in ((Get-ADGroupMember $group.Group -Server $Domaincontroller).SamAccountName))
    {
        $MBX2change +=  get-mailbox $domain$sam -DomainController $Domaincontroller -ErrorAction silentlycontinue
    }

    if($MBX2change.SamAccountName -eq $null)
    {            
        $MBX2change = New-Object –TypeName PSCustomObject
        Add-Member -InputObject $MBX2change -MemberType NoteProperty -Name SamAccountName -Value "nomembers_in_AD_Group_for_custom_quota"
    }


    $changeitems = (Compare-Object -ReferenceObject $changelist.samaccountname -DifferenceObject $MBX2change.samaccountname | ?{$_.SideIndicator -eq "=>"})

    if($changeitems -ne $null)
    {
        foreach($changeitem in $changeitems)
        {
            Set-Mailbox -DomainController $Domaincontroller -IssueWarningQuota $group.Warning -ProhibitSendQuota $group.Send -ProhibitSendReceiveQuota $group.Receive -UseDatabaseQuotaDefaults $false -Identity $domain$($changeitem.InputObject)
            $message = "Set CustomLimits to User " + $changeitem.InputObject + " from group " + $group.group + "."
            Write-Log -Message $message -Path $logfile
        }
    }

    $mbx2change | select samaccountname | export-Csv -Delimiter "|" -Encoding Unicode -Path $changefile

}

$logfiles = Get-ChildItem -Path . -Filter "log*.txt" | ? {$_.creationtime -lt ((Get-Date).adddays(-$LogFileAge))}

foreach($file in $logfiles)
{
	Remove-Item $file.FullName -force
}

Write-Log -Message "script ends here......." -Path $logfile


