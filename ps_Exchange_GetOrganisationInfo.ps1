# Exchange and AD Data extraction

<#

.SYNOPSIS
The puropse of this script is to collect information about the local Active Directory domain and Exchange Organization. 

** It must run on an Exchange server in your organization with Exchange Organization Administrator and Domain Admin rights

.DESCRIPTION

The following data is collected during the operation of this script

AD Domain Naming Information
Domain Controller information
Active Directory Replication subnets
Exchange Server data
Exchange Database Availability Groups
Exchange databases name and state
Summary of mailbox types within Exchange
User mailbox data
Equipment mailbox data
Room mailbox data
Linked mailbox data
Shared mailbox data
Mail contacts data
Distribution group data
Dynamic distribution group data


.EXAMPLE
.\ps_Exchange_GetOrganisationInfo.ps1

.NOTES
Any issues; contact Martin Vogwell at martin.vogwell@ultra-electronics.com

Version history:

Version 0.1 - Development only
Version 1.0 - Beta release
Version 1.1 - Fixed bug with Exchange database enumeration, 
			  Fixed bug with mailboxstatistics using alias instead of Guid, 
			  Added data from Exchange DAGs
			  Changed the output file name to ExchangeDataExtraction.json (old name was confusing)

.LINK
None yet.

#>


[CmdletBinding()]
param ()

$ErrorActionPreference = "Stop"


Function GetMailboxData() {
	param (
		[Parameter(Mandatory=$True)]$arrMbx,
		[Parameter(Mandatory=$True)]$objADData
	)
	
	$arrMailboxData = @()

	Foreach ($objMbx in $arrMbx) {
		$arrEmailAddresses = @()
		
		$objMbx.EmailAddresses | Foreach {
		   $objEmailAddress = New-Object -TypeName PSCustomObject
		   $objEmailAddress | Add-Member -MemberType NoteProperty -Name "Address" -Value $_.AddressString
		   $objEmailAddress | Add-Member -MemberType NoteProperty -Name "PrefixType" -Value $_.PrefixString

		   $arrEmailAddresses += $objEmailAddress
		}

		
		# Remove the email addresses and add the correctly formatted ones
		If (!($null -eq $arrEmailAddresses)) {

			$objMbx = $objMbx | Select-Object -Property * -ExcludeProperty EmailAddresses

			$objMbx | Add-Member -MemberType NoteProperty -Name EmailAddresses -value $arrEmailAddresses
		}

		Try {
			$objRtn = $objADData | Where-Object {$_.ObjectGuid -eq $objMbx.Guid}
		}
		Catch {
			$objRtn = $Null
		}

		Try {
			# Important that this uses objMbx.Guid...tostring() because there may be duplicate alias values in Exchange (with different cases until you try to call a cmdlet using it)
			$mbxStat = Get-MailboxStatistics ($objMbx.Guid).toString() -DomainController $sDC -WarningAction:"SilentlyContinue" | Add-Member -MemberType ScriptProperty -Name TotalItemSizeinMB -Value {$this.totalitemsize.value.ToMB()} -PassThru | Add-Member -MemberType ScriptProperty -Name TotalItemSizeinGB -Value {$this.totalitemsize.value.ToGB()} -PassThru |select TotalItemSizeinMB, TotalItemSizeinGB, lastlogontime, itemcount, databasename 
			
			$objRtn | Add-Member -MemberType NoteProperty -Name "MbxStats" -Value $mbxStat
		}
		Catch {
			$mbxStat = $null
		}

		# Attempt to extract user AD details
		If (!($null -eq $objRtn)) {
			$objRtn | Add-Member -MemberType NoteProperty -Name "DataSourceDomainController" -Value $objRtn.PSComputerName
		
			$objRtn = $objRtn | select-object -Property * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
			
		
			$objMbx | Add-Member -MemberType NoteProperty -Name "ADDetailExtracted" -Value $True

			($objRtn | Get-Member -MemberType NoteProperty).Name | foreach { $objMbx | Add-Member -MemberType NoteProperty -Name $_ -Value $objRtn.$_ }
		}
		Else {
			$objMbx | Add-Member -MemberType NoteProperty -Name "ADDetailExtracted" -Value $False
		}
		
		$arrMailboxData += $objMbx
	}
	
	Return $arrMailboxData
}



#@# Main
Write-Host "`n`nExchange Organization Data Extractor" -ForegroundColor Green
Write-Host "MVogwell - Aug 2020 - Version 1.1`n" -ForegroundColor Green

# Create startup variables
$bStartupSuccess = $True
$objResults = New-Object PSCustomObject

# Get the local domain controller to use for this script
$sDC = ($env:logonserver).substring(2)

#Setup the top level data
# Computername, Timestamp, LocalDC, Domain Netbios, Domain FQDN


# Set the output file path and check the file can be created 
$sOutputFile = Split-Path $MyInvocation.MyCommand.Path -Parent
$sOutputFile = $sOutputFile + "\ExchangeDataExtraction.json"

Try {
	New-Item $sOutputFile -ItemType File -Force | Out-Null
	
	Write-Host "Data output file: $sOutputFile" -ForegroundColor Yellow
}
Catch {
	$sErrorSysMsg = "Error creating output file: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		
	Write-Host $sErrorSysMsg -ForegroundColor Red
	Write-Host "`n`nIt will not be possible to continue...`n`n" -ForegroundColor Red
	
	$bStartupSuccess = $False
}

# Test that the Exchange cmdlets are available
Try {
	Write-Host "Checking availability of Exchange cmdlets: " -ForegroundColor Yellow -NoNewLine
	
	Get-Mailbox -DomainController $sDC | Select-Object -First 1 | Out-Null
	
	Write-Host "Success" -ForegroundColor Green
}
Catch {
	Write-Host "Failed" -ForegroundColor Red
	
	Write-Host "`n`nYou must run this script on an Exchange server with an administrator user`n`n" -ForegroundColor Red

	$bStartupSuccess = $False
}


# Extract the AD data - discover whether the AD cmdlets are available on this local Exchange Server
If ($bStartupSuccess -eq $True) {
	Try {
		$bLocalADCommandsAvailable = $True
		Import-Module ActiveDirectory
	}
	Catch {
		$bLocalADCommandsAvailable = $False
	}

	Try {
		If ($bLocalADCommandsAvailable -eq $True) {		# Import AD worked so get the AD detail from local commands
			
			$objADData = Get-ADUser -Filter * -properties ObjectGuid, City, Company, Country, Department, Description, Enabled, LastLogonDate, PostalCode, State, StreetAddress, Title -Server $sDC -WarningAction:"SilentlyContinue" | select-object ObjectGUID, City, Company, Country, Department, Description, Enabled, LastLogonDate, PostalCode, State, StreetAddress, Title
		}
		Else {	# Attempt to use invoke command against the logon DC to get the AD details

			$objADData = Invoke-Command -ScriptBlock { Get-ADUser -Filter * -properties ObjectGuid, City, Company, Country, Department, Description, Enabled, LastLogonDate, PostalCode, State, StreetAddress, Title -WarningAction:"SilentlyContinue" | select-object ObjectGUID, City, Company, Country, Department, Description, Enabled, LastLogonDate, PostalCode, State, StreetAddress, Title } -ComputerName $sDC
		}
	}
	Catch {
		$sErrorSysMsg = "Error obtaining AD data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		
		Write-Host $sErrorSysMsg -ForegroundColor Red
		Write-Host "`n`nIt will not be possible to continue...`n`n" -ForegroundColor Red
		
		$bStartupSuccess = $False
		$objADData = $null
	}


	# Confirm that AD data is available
	If ($bStartupSuccess -eq $True) {
		Write-Host "Checking access to Active Directory data: " -ForegroundColor Yellow -NoNewLine
	
		# Handle if no AD data was returned
		If ($null -eq $objADData) {
			Write-Host "Failed`n" -ForegroundColor Red
			Write-Host "It has not been possible to extract the Active Directory information from local AD commands or from the local DC. Please contact support`n`n" -ForegroundColor Red
			
			$bStartupSuccess = $False
		}
		Else {
			If (($objADData | Measure-Object).Count -eq 0 ) {
				Write-Host "Failed`n" -ForegroundColor Red
				Write-Host "It has not been possible to extract the Active Directory information from local AD commands or from the local DC. Please contact support`n`n" -ForegroundColor Red
				
				$bStartupSuccess = $False
			}
			Else {
				Write-Host "Success" -ForegroundColor Green
			}
		}
	}
}

#@# Process section - only continue if AD is available
If ($bStartupSuccess -eq $True) { 
	Write-Host "`nExtracting data:" -ForegroundColor Yellow

	Try {
		Write-Host "`t... AD Domain Data" -ForegroundColor Cyan
	
		If ($bLocalADCommandsAvailable -eq $True) {
			$objADDomainData = Get-ADDomain -Server $sDC
				
			$arrDCData = @()

			$objADDomainData.ReplicaDirectoryServers | Foreach {
				$objDCData = Get-ADDomainController $_ -Server $sDC | Select-Object Name, ComputerObjectDN, Enabled, IPv4Address, IsGlobalCatalog, LdapPort, OperatingSystem, OperationMasterRoles
				$arrFSMO = @()
				$objDCData.OperationMasterRoles | ForEach-Object { $arrFSMO += $($_.toString()) }
				$objDCData = $objDCData | Select -Property * -ExcludeProperty OperationMasterRoles
				$objDCData | Add-Member -MemberType NoteProperty -Name "OperationMasterRoles" -Value $arrFSMO
				
				$arrDCData += $objDCData
			}
			
		
		}
		Else {
			$objADDomainData = Invoke-Command -ComputerName $sDC -ScriptBlock { Get-ADDomain }
			
			$arrDCData = @()
			
			$objADDomainData.ReplicaDirectoryServers | Foreach {
				$objDCData = Invoke-Command -ComputerName $sDC -ScriptBlock {Get-ADDomainController | Select-Object Name, ComputerObjectDN, Enabled, IPv4Address, IsGlobalCatalog, LdapPort, OperatingSystem, OperationMasterRoles}
				$arrFSMO = @()
				$objDCData.OperationMasterRoles | ForEach-Object { $arrFSMO += $($_.toString()) }
				$objDCData = $objDCData | Select -Property * -ExcludeProperty OperationMasterRoles
				$objDCData | Add-Member -MemberType NoteProperty -Name "OperationMasterRoles" -Value $arrFSMO
				
				$arrDCData += $objDCData
			}
		}

	
		# Export AD Domain Data
		If ($null -eq $objADDomainData) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ADDomain" -Value "Unable to get data"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "ADDomain" -Value $objADDomainData
		}
		
		# Export Domain Controller data
		If ($null -eq $arrDCData) {
			$objResults | Add-Member -MemberType NoteProperty -Name "DomainControllers" -Value "Unable to get data"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "DomainControllers" -Value $arrDCData
		}
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ADDomain" -Value $sErrorSysMsg	
	}
	

	# Attempt to extract AD replication subnets - this is required to be in a different section as not all versions of AD PS module support it
	Try {
		Write-Host "`t... AD Replication Subnets" -ForegroundColor Cyan
	
		If ($bLocalADCommandsAvailable -eq $True) {
			$objADReplicationSubnet = Get-ADReplicationSubnet -Filter * |Select-Object Site,Name
		}
		Else {
			$objADReplicationSubnet = Invoke-Command -ComputerName $sDC -ScriptBlock {Get-ADReplicationSubnet -Filter * |Select-Object Site,Name}		
		}	
		
		# Export AD Replication Subnets
		If ($null -eq $objADReplicationSubnet) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ADReplicationSubnets" -Value "Unable to get data"
		}
		Else {
			$objADReplicationSubnet = $objADReplicationSubnet | Select-Object -Property * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName
		
			$objResults | Add-Member -MemberType NoteProperty -Name "ADReplicationSubnets" -Value $objADReplicationSubnet
		}		
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ADReplicationSubnets" -Value $sErrorSysMsg	
	}

	
	# Get Details of the Exchange servers - Exchange version info available from https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
	Try {
		Write-Host "`t... Exchange system data" -ForegroundColor Cyan
	
		$arrExchServer = Get-ExchangeServer -DomainController $sDC | Select-Object Name, Fqdn, DistinguishedName, DataPath, Domain, Edition, ExchangeLegacyDN, @{n="Site"; e={$_.Site.Rdn.EscapedName}}, @{n="ExchangeVersion"; e={$_.ExchangeVersion.toString()}}, @{n="AdminDisplayVersion"; e={$_.AdminDisplayVersion.toString()}}

		# Export Domain Controller data
		If ($null -eq $arrExchServer) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeServers" -Value "Unable to get data"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeServers" -Value $arrExchServer
		}		
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeServers" -Value $sErrorSysMsg	
	}
	
	
	# Get Exchange Email Address Policy
	Try {
		Write-Host "`t... Exchange Email Address Policy" -ForegroundColor Cyan
	
		$objExchEmailPol = Get-EmailAddressPolicy | Select-Object @{n='Identity';e={$_.Identity.Name}}, RecipientFilter, @{n='IncludedRecipients'; e={$_.IncludedRecipients.toString()}}, EnabledPrimarySMTPAddressTemplate, @{n='Priority'; e={$_.Priority.toString()}}, EnabledEmailAddressTemplates
		
		# Export Domain Controller data
		If ($null -eq $objExchEmailPol) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeEmailAddrPolicy" -Value "No data returned"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeEmailAddrPolicy" -Value $objExchEmailPol
		}		
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeEmailAddrPolicy" -Value $sErrorSysMsg	
	}

	
	# Get Exchange domains
	Try {
		Write-Host "`t... Exchange domains" -ForegroundColor Cyan
	
		$objExchDomains = Get-AcceptedDomain | Select-Object @{n='Identity';e={$_.Identity.Name}}, @{n='DomainName';e={$_.DomainName.Address}}, MatchSubDomains, @{n='DomainType';e={$_.DomainType.tostring()}}, Default
		
		# Export Domain Controller data
		If ($null -eq $objExchDomains) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeAcceptedDomains" -Value "No data returned"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeAcceptedDomains" -Value $objExchDomains
		}		
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeAcceptedDomains" -Value $sErrorSysMsg	
	}
	
	
	# Get information about the Exchange Database availability groups
	Try {
		$objExchDAGs = Get-DatabaseAvailabilityGroup | Select-Object Name, @{n='Servers'; e={$_.Servers.Rdn.EscapedName -Join(",") }}, @{n='WitnessServer'; e={$_.WitnessServer.Fqdn}}, @{n='WitnessDirectory'; e={$_.WitnessDirectory.PathName}}
		
		# Export Domain Controller data
		If ($null -eq $objExchDAGs) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDAGs" -Value "No data returned"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDAGs" -Value $objExchDAGs
		}		
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDAGs" -Value $sErrorSysMsg	
	}
	
	
	# Get data about the Exchange databases
	Try {
		Write-Host "`t... Exchange databases" -ForegroundColor Cyan
		If ($null -eq $arrExchServer) {
			$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDatabases" -Value "No Exchange Server Data available"
		}
		Else {
			$arrMbxDbCopy = @()
		
			Foreach ($sExchServer in $arrExchServer.Name) {
				Write-Verbose "$sExchServer"

				# Note: DatabaseVolumeMountPoint is reduced to the first two characters because json parsing fails if the value has a trailing \ character
				# Note: Get-MailboxDatabaseCopyStatus has to have -erroraction set to silently continue as some DB copies may fail which would prevent recording of any
				 $arrMbxDbCopy += Get-MailboxDatabaseCopyStatus -Server $sExchServer -DomainController $sDC -WarningAction:"SilentlyContinue"  -ErrorAction "SilentlyContinue" | Select-Object @{n="Identity"; e={$_.Identity.DistinguishedName}}, DatabaseName, Status, MailboxServer, ActiveDatabaseCopy, @{n='DatabaseVolumeMountPoint'; e= {($_.DatabaseVolumeMountPoint).Substring(0,2)}}, DiskFreeSpacePercent
			}
			
			$arrMbxDbCopy = $arrMbxDbCopy | Sort-Object DatabaseName
			
			If ($null -eq $arrMbxDbCopy) {
				$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDatabases" -Value "No databases returned"
			}
			Else {
				$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDatabases" -Value $arrMbxDbCopy
			}
		}
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "ExchangeDatabases" -Value $sErrorSysMsg	
	}
	
	# Get a summary of the types of mailboxes
	Try {
		Write-Host "`t... Mailbox summary data" -ForegroundColor Cyan
	
		$arrMbxTypes = Get-Mailbox -ResultSize Unlimited -DomainController $sDC -WarningAction:"SilentlyContinue" | Group-Object RecipientTypeDetails | Select-Object Count, Name

		$objMbxTypes = New-Object PSCustomObject
			
		# Loop through each mailbox type discovered and tidy up the results into an object to get returned to the overall results
		$arrMbxTypes | ForEach-Object {
			$objMbxTypes | Add-Member -MemberType NoteProperty -Name $($_.Name) -Value $($_.Count)
		}
		
		$objResults | Add-Member -MemberType NoteProperty -Name "MailboxSummary" -Value $objMbxTypes
	}
	Catch {
		$objResults | Add-Member -MemberType NoteProperty -Name "MailboxSummary" -Value "No Data Disovered"
	}

	
	# Get mailbox details (user mailboxes)
	Try {
		Write-Host "`t... User mailbox data" -ForegroundColor Cyan
		
		$arrMbx = Get-Mailbox -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"} | Select-Object Name, Alias, DisplayName, DistinguishedName, SamAccountName, userPrincipalName, Guid, IsLinked, IsMailboxEnabled, IsShared, LegacyExchangeDN, LinkedMasterAccount, @{n='MaxReceiveSize' ;e={$_.MaxReceiveSize.Value}}, @{n='MaxSendSize' ;e={$_.MaxSendSize.Value}}, Office, OrganizationalUnit, RetentionPolicy, ArchiveDatabase, AccountDisabled, @{n='DatabaseName' ;e={$_.Database.Name}}, @{n='DatabaseDN' ;e={$_.Database.DistinguishedName}}, @{n='DatabaseServerName' ;e={$_.ServerName}}, DeliverToMailboxAndForward, UMEnabled, @{n='WhenCreatedUTC' ; e={$_.WhenCreatedUTC.toString()}}, EmailAddresses, @{n='WindowsEmailAddress'; e={$_.WindowsEmailAddress.Address}}, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}, EmailAddressPolicyEnabled
		$arrMbxResults = GetMailboxData -arrMbx $arrMbx -objADData $objADData
		$objResults | Add-Member -MemberType NoteProperty -Name "UserMailboxes" -Value $arrMbxResults
	}
	Catch {
		$objResults | Add-Member -MemberType NoteProperty -Name "UserMailboxes" -Value "No Data Disovered"
	}
	
	# Get mailbox details (Equipment mailboxes)
	Try {
		Write-Host "`t... Equipment mailbox data" -ForegroundColor Cyan
	
		$arrMbx = Get-Mailbox -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Where-Object {$_.RecipientTypeDetails -eq "EquipmentMailbox"} | Select-Object Name, Alias, DisplayName, DistinguishedName, SamAccountName, userPrincipalName, Guid, IsLinked, IsMailboxEnabled, IsShared, LegacyExchangeDN, LinkedMasterAccount, @{n='MaxReceiveSize' ;e={$_.MaxReceiveSize.Value}}, @{n='MaxSendSize' ;e={$_.MaxSendSize.Value}}, Office, OrganizationalUnit, RetentionPolicy, ArchiveDatabase, AccountDisabled, @{n='DatabaseName' ;e={$_.Database.Name}}, @{n='DatabaseDN' ;e={$_.Database.DistinguishedName}}, @{n='DatabaseServerName' ;e={$_.ServerName}}, DeliverToMailboxAndForward, UMEnabled, @{n='WhenCreatedUTC' ; e={$_.WhenCreatedUTC.toString()}}, EmailAddresses, @{n='WindowsEmailAddress'; e={$_.WindowsEmailAddress.Address}}, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}, EmailAddressPolicyEnabled
		$arrMbxResults =  GetMailboxData -arrMbx $arrMbx -objADData $objADData

		If ($null -eq $arrMbxResults) { 
			$objResults | Add-Member -MemberType NoteProperty -Name "EquipmentMailboxes" -Value "No data returned"
		} 
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "EquipmentMailboxes" -Value $arrMbxResults
		}
		
		}
	Catch {
		$objResults | Add-Member -MemberType NoteProperty -Name "EquipmentMailboxes" -Value "No Data Disovered"
	}
	
	
	# Get mailbox details (Room mailboxes)
	Try {
		Write-Host "`t... Room mailbox data" -ForegroundColor Cyan
		
		$arrMbx = Get-Mailbox -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Where-Object {$_.RecipientTypeDetails -eq "RoomMailbox"} | Select-Object Name, Alias, DisplayName, DistinguishedName, SamAccountName, userPrincipalName, Guid, IsLinked, IsMailboxEnabled, IsShared, LegacyExchangeDN, LinkedMasterAccount, @{n='MaxReceiveSize' ;e={$_.MaxReceiveSize.Value}}, @{n='MaxSendSize' ;e={$_.MaxSendSize.Value}}, Office, OrganizationalUnit, RetentionPolicy, ArchiveDatabase, AccountDisabled, @{n='DatabaseName' ;e={$_.Database.Name}}, @{n='DatabaseDN' ;e={$_.Database.DistinguishedName}}, @{n='DatabaseServerName' ;e={$_.ServerName}}, DeliverToMailboxAndForward, UMEnabled, @{n='WhenCreatedUTC' ; e={$_.WhenCreatedUTC.toString()}}, EmailAddresses, @{n='WindowsEmailAddress'; e={$_.WindowsEmailAddress.Address}}, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}, EmailAddressPolicyEnabled
		$arrMbxResults =  GetMailboxData -arrMbx $arrMbx -objADData $objADData
		
		If ($null -eq $arrMbxResults) { 
			$objResults | Add-Member -MemberType NoteProperty -Name "RoomMailboxes" -Value "No data returned"
		} 
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "RoomMailboxes" -Value $arrMbxResults
		}
	}
	Catch {
		$objResults | Add-Member -MemberType NoteProperty -Name "RoomMailboxes" -Value "No Data Disovered"
	}
	
	
	# Get mailbox details (Linked mailboxes)
	Try {
		Write-Host "`t... Linked mailbox data" -ForegroundColor Cyan
	
		$arrMbx = Get-Mailbox -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Where-Object {$_.RecipientTypeDetails -eq "LinkedMailbox"} | Select-Object Name, Alias, DisplayName, DistinguishedName, SamAccountName, userPrincipalName, Guid, IsLinked, IsMailboxEnabled, IsShared, LegacyExchangeDN, LinkedMasterAccount, @{n='MaxReceiveSize' ;e={$_.MaxReceiveSize.Value}}, @{n='MaxSendSize' ;e={$_.MaxSendSize.Value}}, Office, OrganizationalUnit, RetentionPolicy, ArchiveDatabase, AccountDisabled, @{n='DatabaseName' ;e={$_.Database.Name}}, @{n='DatabaseDN' ;e={$_.Database.DistinguishedName}}, @{n='DatabaseServerName' ;e={$_.ServerName}}, DeliverToMailboxAndForward, UMEnabled, @{n='WhenCreatedUTC' ; e={$_.WhenCreatedUTC.toString()}}, EmailAddresses, @{n='WindowsEmailAddress'; e={$_.WindowsEmailAddress.Address}}, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}, EmailAddressPolicyEnabled
		$arrMbxResults =  GetMailboxData -arrMbx $arrMbx -objADData $objADData
		
		If ($null -eq $arrMbxResults) { 
			$objResults | Add-Member -MemberType NoteProperty -Name "LinkedMailbox" -Value "No data returned"
		} 
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "LinkedMailbox" -Value $arrMbxResults
		}
	}
	Catch {
		$objResults | Add-Member -MemberType NoteProperty -Name "LinkedMailbox" -Value "No Data Disovered"
	}


	# Get mailbox details (Shared mailboxes)
	Try {
		Write-Host "`t... Shared mailbox data" -ForegroundColor Cyan
		
		$arrMbx = Get-Mailbox -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"} | Select-Object Name, Alias, DisplayName, DistinguishedName, SamAccountName, userPrincipalName, Guid, IsLinked, IsMailboxEnabled, IsShared, LegacyExchangeDN, LinkedMasterAccount, @{n='MaxReceiveSize' ;e={$_.MaxReceiveSize.Value}}, @{n='MaxSendSize' ;e={$_.MaxSendSize.Value}}, Office, OrganizationalUnit, RetentionPolicy, ArchiveDatabase, AccountDisabled, @{n='DatabaseName' ;e={$_.Database.Name}}, @{n='DatabaseDN' ;e={$_.Database.DistinguishedName}}, @{n='DatabaseServerName' ;e={$_.ServerName}}, DeliverToMailboxAndForward, UMEnabled, @{n='WhenCreatedUTC' ; e={$_.WhenCreatedUTC.toString()}}, EmailAddresses, @{n='WindowsEmailAddress'; e={$_.WindowsEmailAddress.Address}}, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}, EmailAddressPolicyEnabled
		$arrMbxResults =  GetMailboxData -arrMbx $arrMbx -objADData $objADData
		
		If ($null -eq $arrMbxResults) { 
			$objResults | Add-Member -MemberType NoteProperty -Name "SharedMailbox" -Value "No data returned"
		} 
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "SharedMailbox" -Value $arrMbxResults
		}
	}
	Catch {
		$objResults | Add-Member -MemberType NoteProperty -Name "SharedMailbox" -Value "No Data Disovered"
	}


	# Get mailcontacts data (Excluding UEGAL)
	Try {
		Write-Host "`t... Mail contact data" -ForegroundColor Cyan
	
		$arrContacts = Get-MailContact -Filter "(CustomAttribute5 -ne 'UEGAL')" -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Select-Object Alias, Name, DisplayName, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}, @{n='ExternalEmailAddress'; e={$_.ExternalEmailAddress.Address}}, EmailAddresses, HiddenFromAddressListsEnabled, OrganizationalUnit
		
		If ($null -eq $arrContacts) {
			$objResults | Add-Member -MemberType NoteProperty -Name "MailContacts" -Value "No contacts discovered"
		}
		Else {
			$arrContactsParsed = @()
		
			Foreach ($objContact in $arrContacts) {
				$arrEmailAddresses = @()
				$arrContacts.EmailAddresses | ForEach-Object {
					$arrEmailAddresses += $_.ProxyAddressString
				}
				
				$objContact = $objContact | Select-Object -Property * -ExcludeProperty EmailAddresses
				$objContact | Add-Member -MemberType NoteProperty -Name EmailAddresses -Value $arrEmailAddresses
				
				$arrContactsParsed += $objContact
			}
			
			# Remove the old array as it should no longer be required
			$arrContacts = $Null
						
			$objResults | Add-Member -MemberType NoteProperty -Name "MailContacts" -Value $arrContactsParsed
		}
	}
	Catch {
		$sErrorSysMsg = "Error extracting data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "MailContacts" -Value $sErrorSysMsg
	}


	# Get groups and group membership
	Try {
		Write-Host "`t... Distribution group data" -ForegroundColor Cyan
	
		$arrGroups = Get-DistributionGroup -DomainController $sDC -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Select-Object Name, Alias, DisplayName, @{n='GroupType'; e={$_.GroupType.ToString()}}, SamAccountName, OrganizationalUnit, EmailAddresses, HiddenFromAddressListsEnabled, LegacyExchangeDN, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}} 

		If ($null -eq $arrGroups) {
			$objResults | Add-Member -MemberType NoteProperty -Name "DistributionGroupData" -Value "No Distribution Group data discovered"
		} 
		Else {
			$arrGroupsOutput = @()
			
			Foreach ($objGroup in $arrGroups) {
				$arrGrpMembers = Get-DistributionGroupMember $objGroup.Alias -ResultSize Unlimited -WarningAction:"SilentlyContinue" | Select-Object Alias, @{n='RecipientType'; e={$_.RecipientType.ToString()}}, RecipientTypeDetails
			
				if ($null -eq $arrGrpMembers) {
					$objGroup | Add-Member -MemberType NoteProperty -Name "Members" -Value "Null returned"
				}
				ElseIf($arrGrpMembers.Count -le 0) {
					$objGroup | Add-Member -MemberType NoteProperty -Name "Members" -Value "No members"
				}
				Else {
					$objGroup | Add-Member -MemberType NoteProperty -Name "Members" -Value $arrGrpMembers
				}
				
				$arrGroupsOutput += $objGroup
			}
			
			$objResults | Add-Member -MemberType NoteProperty -Name "DistributionGroupData" -Value $arrGroupsOutput
		}
	}
	Catch {
		$sErrorSysMsg = "Error extracting group data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "DistributionGroupData" -Value $sErrorSysMsg
	} # End of: Get groups and group membership
	

	# Get Dynamic Distribution Group data
	Try {
		Write-Host "`t... Dynamic distribution group data" -ForegroundColor Cyan
	
		$arrDynDistGrps = Get-DynamicDistributionGroup -ResultSize Unlimited -DomainController $sDC -WarningAction:"SilentlyContinue" | Select-Object Alias, Name, DisplayName, RequireSenderAuthenticationEnabled, RecipientFilter, @{n='IncludedRecipients'; e={$_.IncludedRecipients.tostring()}}, @{n='PrimarySmtpAddress'; e={$_.PrimarySmtpAddress.Address}}
		
		If ($null -eq $arrDynDistGrps) {
			$objResults | Add-Member -MemberType NoteProperty -Name "DynamicDistributionGroupData" -Value "None discovered"
		}
		Else {
			$objResults | Add-Member -MemberType NoteProperty -Name "DynamicDistributionGroupData" -Value $arrDynDistGrps
		}
	}
	Catch {
		$sErrorSysMsg = "Error extracting dynamic dist group data: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
		$objResults | Add-Member -MemberType NoteProperty -Name "DynamicDistributionGroupData" -Value $sErrorSysMsg
	}

	Try {
		$bErrorExporting = $False		
	
		$objResults | ConvertTo-Json -Depth 4 | Out-File $sOutputFile
	
		Write-Host "`nOutput file $sOutputFile created successfully (JSON)" -ForegroundColor Green
		
		Write-Host "`nFinished! Please email the output file back to your Group IT contact`n`n" -ForegroundColor Green
	}
	Catch {
		$bErrorExporting = $True
	}

	# Attempt to export as Json using the compress option
	If ($bErrorExporting -eq $True) {
		Try {
			$bErrorExporting = $False		
		
			$objResults | ConvertTo-Json -Depth 4 -Compress | Out-File $sOutputFile
		
			Write-Host "`nOutput file $sOutputFile created successfully (JSON Compressed)" -ForegroundColor Green
			
			Write-Host "`nFinished! Please email the output file back to your Group IT contact`n`n" -ForegroundColor Green
		}
		Catch {
			$bErrorExporting = $True
		}
	}
	
	If ($bErrorExporting -eq $True) {
		Try {
			$objResults | Export-CliXML $sOutputFile -Depth 4

			Write-Host "`nOutput file $sOutputFile created successfully (XML)" -ForegroundColor Green
			
			Write-Host "`nFinished! Please email the output file back to your Group IT contact`n`n" -ForegroundColor Green
		}
		Catch {
			$sErrorSysMsg = "`nError saving data to file: " + (($Error[0].exception).toString()).replace("`r"," ").replace("`n"," ")
			Write-Host $sErrorSysMsg -ForegroundColor Red
		
			Write-Host "`nError completing the data extraction`n`n" -ForegroundColor Red
		}
	}
}



