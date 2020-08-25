# ps_Exchange_GetOrganisationInfo.ps1

## About

The puropse of this script is to collect information about the local Active Directory domain and Exchange Organization. 

** It must run on an Exchange server in your organization with Exchange Organization Administrator and Domain Admin rights

The following data is collected during the operation of this script

* AD Domain Naming Information
* Domain Controller information
* Active Directory Replication subnets
* Exchange Server data
* Exchange Database Availability Groups
* Exchange databases name and state
* Summary of mailbox types within Exchange
* User mailbox data
* Equipment mailbox data
* Room mailbox data
* Linked mailbox data
* Shared mailbox data
* Mail contacts data
* Distribution group data
* Dynamic distribution group data

## Requirements

This script must be run in the following environment:

* Run within Exchange Management Shell on an Exchange server
    * Exchange 2013 or later is preferred. Exchange 2010 with Powershell version will not return all of the possible data as not all cmdlets are supported.
* Run with a user who is a member of the Exchange Organization Administrators
* If the Powershell Active Directory module is available on the local machine it will use this. Otherwise a domain controller must be available that has the firewall rules enabled to allow WinRM traffic.
* To be able to get the domain information it is best if the user account is a member of the Domain Admins group.


## Output

The output of this file is a Json file (ExchangeDataExtraction.json) which is saved to the same folder as the script is run from. If the conversion to json fails it will attempt to convert to json compressed and if that in turn fails it will attempt to export to XMLCLI. During the output the data type will be displayed.


## Troubleshooting 

If the script will not run then it could be for the following reasons:

### Execution Policy
First try unblocking the script with: `unblock-file "file path"` and then re-running the script

If this doesn't work then set the execution policy to RemoteSigned:

```powershell
Get-ExecutionPolicy     # Make a note of the current execution policy to be able to change it back again
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

# And then once you've finished running the script:
Set-ExecutionPolicy -ExecutionPolicy _"Whatever it was before"_
```


## Reading the results in PowerShell

You can use the following commands to import the data back into PowerShell for analysis:

### Json or Json compressed data
$data = Get-Content "Full path of the data file" | ConvertFrom-Json

### XML data
$data = Import-CliXML "Full path of the data file"

### Listing the top level data once the data has been imported:

($data | Get-Member -MemberType NoteProperty).Name


### Notes
No responsibility is taken for you running this script or any outcomes.