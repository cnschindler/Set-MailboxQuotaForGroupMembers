#############################
#
# Set-NewQuota.ps1
# Resets Users to default Quota if they are
# not Member of a Custom Quota Group
#
# Sets Quota of Groupmembers to predefined Values
#
#############################
# Domaincontroller, Group and Limit Declaration
#############################
# (<groupname>,<Warning>,<sendlimit>,<receivelimit>)

$Domaincontroller="dc03.noehw.local"
$domain = "noehw\"
$LogFileAge = 14

$mygroupandlimits = @(
             ("MBX2016_Quota_1GB","750MB","1024MB","1250MB"),
             ("MBX2016_Quota_4GB","3584MB","4096MB","4403MB"),
             ("MBX2016_Quota_8GB","7168MB","8192MB","8806MB"),
             ("MBX2016_Quota_16GB","14336MB","16384MB","17612MB")
   )


##############################
# Script starts here
##############################

# 1.0 Check Powershell Version
if ((Get-Host).Version.Major -eq 1)
{
	throw "Powershell Version 1 not supported";
}

# Check AD Module, attempt to load
if (!(Get-Command Get-ADGroupMember -ErrorAction SilentlyContinue))
{
    Import-Module ActiveDirectory
    if(!(Get-Module ActiveDirectory))
    {
        throw "Active Directory Module cannot be loaded"
    }
}
    

# Check Exchange Management Shell, attempt to load
if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
	
	if (Test-Path "C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1")
	{
		. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
		Connect-ExchangeServer -auto
	} elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1") {
		Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
		.'C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1'
	} elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1") {
		. 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
		Connect-ExchangeServer -auto
	}else {
		throw "Exchange Management Shell cannot be loaded"
	}
}

#Push-Location (Split-Path $script:MyInvocation.MyCommand.Path)


. C:\Scripts\MailboxQuota\Function-Write-Log.ps1

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

#write-host $GroupLimits

######################
# Routines
######################
# remove users only
######################

$logfile = ".\Log\LOG_"+(Get-Date -Format ddMMyyyyHHmm)+".txt"
#
#Write-Log -Message "**********************************************************************************" -Path $logfile
#Write-Log -Message "Remove Users without Mailbox from Groups...." -Path $logfile
#
#foreach($group in $GroupLimits)
#{
#    $quotagroup2members = $quotagroup2members = Get-ADGroupMember $group.group -Server $Domaincontroller| Get-aduser -Properties msExchRecipientTypeDetails
#    $group2membersfiltered = ($quotagroup2members | ?{$_.msExchRecipientTypeDetails -eq $null})
#    #$group2membersfiltered | ?{$_.recipienttype -ne "usermailbox"}
#
#    foreach($gmf in $group2membersfiltered)
#    {
#	    Remove-ADGroupMember -Identity $group.group -Member $gmf.SamAccountName -Server $Domaincontroller -Confirm:$false
#        $message = "Removing distribution group member " + $gmf.SamAccountName + " from distribution group " + $group.group + "."
#        Write-Log -Message $message -Path $logfile
#    }
#
#}

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

$Filter = @(
    '(RecipientType -eq "UserMailbox")'
    '(Displayname -notlike "Discovery Search Mailbox*")'
    '(Displayname -ne "Personal 1")'
    '(Displayname -ne "Personal 2")'
    '(Displayname -ne "Goll Gabriela")'
    '(Displayname -ne "Gleiß Regina")'
    '(UseDatabaseQuotaDefaults -eq $false)'
) -join ' -and '

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


