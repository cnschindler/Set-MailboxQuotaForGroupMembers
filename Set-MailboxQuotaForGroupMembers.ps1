########################################################################
#
# Set-MailboxQuotasForGroupMembers.ps1
# 
# Requirements: Active Directory PowerShell Module and Exchange Management Shell
#
# Version: 1.0
# Author: Christian Schindler, NTx BackOffice Consulting Group GmbH
# Contact: christian.schindler@ntx.at
# Date: 2025-07-07
# This script is provided as-is without any warranties. Use at your own risk.
# License: GNU General Public License v3.0 (GPL-3.0)
#
########################################################################

<#
.SYNOPSIS
    Sets Mailbox-Storage-Quotas for members of a group to predefined Values
.DESCRIPTION
    This script sets mailbox quotas for members of specified Active Directory groups based on a JSON configuration file.
    The JSON configuration file defines the groups and their corresponding mailbox quota settings such as IssueWarningQuota,
    ProhibitSendQuota, and ProhibitSendReceiveQuota.

    The script retrieves group members, checks if they have mailboxes, and applies the specified quota settings to each mailbox.
    It also includes logging functionality to track the progress and any errors encountered during execution.
.EXAMPLE
    Set-MailboxQuotaForGroupMembers.ps1 -ConfigFile "C:\Path\To\Config.json"

    This example runs the script with a specified configuration file that contains the group names and their quota settings.
.NOTES
    Version: 1.0
.PARAMETER ConfigFile
    Specifies the path to the JSON configuration file that contains the group names and their corresponding mailbox quota settings.
    If not specified, it defaults to "Set-MailboxQuotaForGroupMembers-Config.json" in the script's directory.
#>

# Check if the script is running in PowerShell version 3.0 or higher and if the ActiveDirectory module is available
# Requires statement ensures that the script will not run if the required version or module is not present
#Requires -Version 3.0
#Requires -Module ActiveDirectory

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [System.IO.FileInfo]
    $ConfigFile = (Join-Path -Path $PSScriptRoot -ChildPath "Set-MailboxQuotaForGroupMembers-Config.json")
)

# Global variables
# Define the path for the logfile, using the script name and current date/time for uniqueness
$Scriptname = $MyInvocation.MyCommand.Name.Replace(".ps1", "")
[System.IO.FileInfo]$LogfileFullPath = Join-Path -Path $PSScriptRoot -ChildPath ($Scriptname  + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
$script:NoLogging

# Logging function, used for progress and error logging...
# Uses the globally (script scoped) configured variables 'LogfileFullPath' to identify the logfile and 'NoLogging' to disable it.
function Write-LogFile
{
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )

    # Prefix the string to write with the current Date and Time, add error message if present...
    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : ERROR: {1} Error: {2}" -f [DateTime]::Now, $Message, $ErrorInfo.Exception.Message
    }

    else
    {
        $logLine = "{0:d.M.y H:mm:ss} : INFO: {1}" -f [DateTime]::Now, $Message
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

Function Import-ADModule
{
    $ModuleName = "ActiveDirectory"
    try
    {
        Import-Module -Name $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
        Write-LogFile -Message "ActiveDirectory Module successfully loaded."
    }
    
    catch
    {
        Write-LogFile -Message "ActiveDirectory Module could not be loaded." -ErrorInfo $_
    }
}
function Connect-ExchangeOnPremieses
{
    # Check if a connection to an exchange server exists and connect if necessary...
    if (-NOT (Get-PSSession | Where-Object ConfigurationName -EQ "Microsoft.Exchange"))
    {
        # Test if Exchange Management Shell Module is installed - if not, exit the script
        $EMSModuleFile = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction SilentlyContinue).MsiInstallPath + "bin\RemoteExchange.ps1"
        
        # If the EMS Module wasn't found
        if (-Not (Test-Path $EMSModuleFile))
        {
            # Write Error end exit the script
            $ErrorMessage = "Exchange Management Shell Module not found on this computer. Please run this script on a computer with Exchange Management Tools installed!"
            Write-LogFile -Message $ErrorMessage
            Exit
        }

        else
        {
            # Load Exchange Management Shell
            try
            {
                # Dot source the EMS Script
                . $($EMSModuleFile) -ErrorAction Stop | Out-Null
                Write-LogFile -Message "Successfully loaded Exchange Management Shell Module."
            }

            catch
            {
                Write-LogFile -Message "Unable to load Exchange Management Shell Module." -ErrorInfo $_
                Exit
            }

            # Connect to Exchange Server
            try
            {
                Connect-ExchangeServer -auto -ClientApplication:ManagementShell -ErrorAction Stop | Out-Null
                Write-LogFile -Message "Successfully connected to Exchange Server."
            }

            catch
            {
                Write-LogFile -Message "Unable to connect to Exchange Server." -ErrorInfo $_
                Exit
            }
        }
    }
}

# Load Active Directory Module
Import-ADModule

# Connect to Exchange Server
Connect-ExchangeOnPremieses

# Import the configuration file
# Check if the configuration file exists
if (-not (Test-Path -Path $ConfigFile -PathType Leaf))
{
    Write-LogFile -Message "Configuration file not found: $ConfigFile"
    Exit
}

# Load the configuration file
Write-LogFile -Message "Loading configuration file: $ConfigFile"
$Config = Get-Content -Path $ConfigFile | ConvertFrom-Json
[string]$Domaincontroller = $Config.Domaincontroller
$Quotas = $Config.Quotas
$LogFileAge = $Config.LogFileAge

# Check if a Domaincontroller was specified in the configuration file, otherwise use the PrimaryDC
if ([System.String]::IsNullOrEmpty($Domaincontroller))
{
    $Domaincontroller = (Get-ADDomainController -Discover -Service "PrimaryDC").HostName
    Write-LogFile -Message "No Domaincontroller specified, using PrimaryDC: $Domaincontroller"
}

else
{
    Write-LogFile -Message "Using Domaincontroller specified in config file: $Domaincontroller"
}

# Check if the Domaincontroller is reachable
Write-LogFile -Message "Checking if Domaincontroller $Domaincontroller is reachable..."
# Using Test-Connection to verify network connectivity to the Domaincontroller
if (-not (Test-Connection -ComputerName $Domaincontroller -Count 1 -ErrorAction SilentlyContinue))
{
    Write-LogFile -Message "Domaincontroller $Domaincontroller is not reachable. Please check the network connection or the Domaincontroller name."
    Exit
}

else
{
    Write-LogFile -Message "Domaincontroller $Domaincontroller is reachable."
}

# Start setting mailbox quotas for group members
Write-LogFile -Message "Starting to set mailbox quotas for group members."

# Iterate through each quota entry defined in the configuration file
foreach($entry in $Quotas)
{
    Write-LogFile -Message "Processing group: $($entry.Name) with settings: IssueWarning=$($entry.Settings.IssueWarning), ProhibitSend=$($entry.Settings.ProhibitSend), ProhibitSendReceive=$($entry.Settings.ProhibitSendReceive)"

    # Get all group members and their mailbox properties
    # Using Get-ADGroupMember to retrieve group members and then filtering for mailboxes
    $Groupmember = foreach ($Member in ((Get-ADGroupMember $entry.Name -Server $Domaincontroller).SamAccountName))
    {
        get-mailbox $Member -DomainController $Domaincontroller -ErrorAction SilentlyContinue
    }

    # If group members were found, set their mailbox quotas
    if ($Groupmember.count -gt 0)
    {
        Write-LogFile -Message "Found $($Groupmember.count) members in group '$($entry.Name)'. Setting mailbox quotas..."
        Write-LogFile -Message "If quota settings are already set to the values specified in the config file, the mailbox will be skipped."
        
        # Iterate through each member of the group
        foreach ($member in $Groupmember)
        {
            if($Entry.Settings.IssueWarning.Replace("MB","") -eq $Member.IssueWarningQuota.value.ToMB() -and
               $Entry.Settings.ProhibitSend.Replace("MB","") -eq $Member.ProhibitSendQuota.value.ToMB() -and
               $Entry.Settings.ProhibitSendReceive.Replace("MB","") -eq $Member.ProhibitSendReceiveQuota.value.ToMB())
            {
                #Write-LogFile -Message "Skipping $($member.SamAccountName) as the quota is already set to values specified in config file."
                Continue
            }
            
            Else
            {
                try
                {
                    # Set mailbox quotas for the current member
                    Write-LogFile -Message "Trying to set custom Quota for $($member.SamAccountName)"
                    Set-Mailbox -Identity $member.SamAccountName -IssueWarningQuota $entry.Settings.IssueWarning -ProhibitSendQuota $entry.Settings.ProhibitSend -ProhibitSendReceiveQuota $entry.Settings.ProhibitSendReceive -UseDatabaseQuotaDefaults $false -DomainController $Domaincontroller -ErrorAction Stop
                    Write-LogFile -Message "Successfully set Custom Quota for $($member.SamAccountName) to $($entry.Settings.IssueWarning), $($entry.Settings.ProhibitSend), $($entry.Settings.ProhibitSendReceive)."
                }

                catch
                {
                    Write-LogFile -Message "Error setting Custom Quota for $($member.SamAccountName):" -ErrorInfo $_
                }
            }
        }
    }

    else
    {
        # If no members were found in the group, log a message and continue to the next group
        Write-LogFile -Message "No members found in group '$($entry.Name)'. Skipping."
        Continue
    }
}

Write-LogFile -Message "Finished setting mailbox quotas for group members."

# Clean up old logfiles based on the configured age
# Retrieve logfiles older than the configured age
$logfiles = Get-ChildItem -Path $LogfileFullPath.Directory -Filter "*.log" | Where-Object {$_.creationtime -lt ((Get-Date).adddays(-$LogFileAge))}

# Remove old logfiles
if ($logfiles.Count -gt 0)
{
    foreach($file in $logfiles)
    {
        try
        {
            # Remove the logfile
            Remove-Item $file.FullName -force -ErrorAction Stop -Confirm:$false -WhatIf:$false
            Write-LogFile -Message "Successfully removed logfile: $($file.FullName)"
        }

        catch
        {
            # If an error occurs while removing the logfile, log the error
            Write-LogFile -Message "Error removing logfile: $($file.FullName)" -ErrorInfo $_
        }
    }
}

# End of script
Write-LogFile -Message "Script execution completed."
