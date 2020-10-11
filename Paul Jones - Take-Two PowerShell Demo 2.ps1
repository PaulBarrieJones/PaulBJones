#######################################################################################################################


#### Paul Jones - Take-Two - Script example 2 - Add users to Microsoft Team from group of users in BlackBerry Enterprise
####
#### Specific users (several thousand) were manually joined to a Microsoft Team by the service desk. The turnaround was
#### hundreds at a time, with users leaving at a similar rate. The group was part of a seperate system not connected to
#### Active Directory. With the API of Blackberry I wrote this automation script to run daily to query the BES data and
#### extract user date of the specific group, each member is then added to the Team. Users no longer in the group are 
#### are now also removed from the group. This script has saved hundreds of hours of service desk time. This script was
#### designed and written in approx. 3 hours and runs on a scheduled task.
#### Removal from the Team is handled in a different process and is not required here.
#### Please note, the script has been sanitised and any propriatary information has been renamed for security purposes.


#######################################################################################################################



# # Requires -Module msonline.psm1  >
# # Requires PowerShell Version 3 or above  >

<#

.SYNOPSIS


.DESCRIPTION

.PARAMETERS

.INPUTS


.OUTPUTS


.NOTES
  Version:          1.0
  Author:           Paul Jones
  Creation Date:    13/12/2019
  Purpose/Change:   Automation
  Change Ref:       CAB000000
  Change Type:      Normal
  
.EXAMPLE
  None

#>


#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Gather initial system varibles at PowerShell instance start
$DefaultVariables = Get-Variable | Select-Object -ExpandProperty Name

# Set Error Verbose and Warning
$ErrorActionPreference = "Continue"
$VerbosePreference = "continue"
$WarningPreference = "continue"
$ExitCode = 0

# Load Required Module/Function Libraries

Try {

    Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\MSOnline\MSOnline.psm1 -ErrorAction Stop

    }

        Catch {

        Clear-Host

        Write-Warning "MSonline.psm1 module not present on system, script cannot continue!"
        Write-Warning "To install modude: Find-Module MSonline | Install-Module"

        BREAK

        }


### Allow unsecure connection to HTTPS (required)

If (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type) {

$certCallback = @"
using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public class ServerCertificateValidationCallback
{
public static void Ignore()
{
if(ServicePointManager.ServerCertificateValidationCallback ==null)
{
ServicePointManager.ServerCertificateValidationCallback += 
delegate
(
Object obj, 
X509Certificate certificate, 
X509Chain chain, 
SslPolicyErrors errors
)
{
return true;
};
}
}
}
"@

Add-Type $certCallback

}

[ServerCertificateValidationCallback]::Ignore()


#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script Version
$ScriptVersion = "1.0"
$ProjectName = "ADD_TO_TEAM"
$ProjectPath = "C:\PowerShell\Projects\$ProjectName"
$LogFilePath = "$ProjectPath\Logs\$ProjectName.log"
$ScriptPath = "$ProjectPath\Scripts\"

# Log File Info

$LogFilePath = "$ProjectPath\Logs\$ProjectName.log"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

### Logging function

function Write-Log {
    
    param (
        
        [Parameter(Mandatory=$False, Position=0)]
        [String]$Entry
    
    )

    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') $Entry" | Out-File -FilePath $LogFilePath -Append
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

### Log Start
Write-Log -Entry "Script DHCP started on $(Get-Date -Format 'dddd, MMMM dd, yyyy')."
Write-Log -Entry "Script Version: $ScriptVersion"
Write-Log -Entry "Script Path: $ScriptPath"

### Connect to BES Service
$UEMHostAndPort = "SERVER:2020"
$TenantGuid = "XXXXXXXX"
$global:UEMHostPortTenantGUIDBaseURL = "https://"+$UEMHostAndPort + "/" + $TenantGuid+"/api/v1"

# Read only BES account credentials
$username = "BESAPI"
$Base64Password = @(000000)

# Create and set the headers for the REST request
$headers = @{}
$headers.Add("Content-Type", "application/vnd.blackberry.authorizationrequest-v1+json")

# Create body for REST call
$body = @{}
$body.add("provider","AD");
$body.add("username", $UserName);
$body.add("password", $Base64Password);
$body.add("domain", "BTP.local");
   
$request = $global:UEMHostPortTenantGUIDBaseURL+"/util/authorization";
$response = Invoke-RESTMethod -Method POST -Headers $headers -Uri $request -Body ($body | ConvertTo-Json);
$global:AuthorizationString = $response

$headers = @{}
$headers.Add("Authorization", $global:AuthorizationString);
$headers.Add("Content-Type", "application/vnd.blackberry.userdetail-v1+json");

# Split results into 5 pages due to limitations within the API
$page1 = Invoke-RestMethod -Method get -uri "$global:UEMHostPortTenantGUIDBaseURL/users?query=groupGuid=00000000000-000000000-0000=1000&offset=0" -Headers $headers
$page2 = Invoke-RestMethod -Method get -uri "$global:UEMHostPortTenantGUIDBaseURL/users?query=groupGuid=00000000000-000000000-0000=1000&offset=1000" -Headers $headers
$page3 = Invoke-RestMethod -Method get -uri "$global:UEMHostPortTenantGUIDBaseURL/users?query=groupGuid=00000000000-000000000-0000=1000&offset=2000" -Headers $headers
$page4 = Invoke-RestMethod -Method get -uri "$global:UEMHostPortTenantGUIDBaseURL/users?query=groupGuid=00000000000-000000000-0000=1000&offset=3000" -Headers $headers
$page5 = Invoke-RestMethod -Method get -uri "$global:UEMHostPortTenantGUIDBaseURL/users?query=groupGuid=00000000000-000000000-0000=1000&offset=4000" -Headers $headers

# Combine the pages of users for next stage
$TotalUsers = $page1.users, $page2.users, $page3.users, $page4.users, $page5.users

# Identify users to be added to Team
$UserstoAdd = $TotalUsers.username
# Get UPN to add the user to Team
$usersforTeam = $TotalUsers.emailaddress

# Create array of users to be added
$UsersAdded = @()

# Add users to Team
Foreach ($User in $UserstoAdd) {

    Try {

    Add-ADGroupMember -Identity "TEAM_SPECIAL" -Members $user -ErrorAction stop
    write-host "Adding: $user" -ForegroundColor Green

    $UsersAdded += $User

    }

    Catch {
    write-host "Exising: $User" -ForegroundColor Yellow

    }

}


# ReadOnly user/service account for connection to tenency, has permission to manage users of TEAM only
# Script is held in secure folder and is only readable by service account on scheduled task

$AdminUser = "TeamAdminAccount"

$ADPassword = "00000000000000000000000000000000000000000000"

$ADKey = @(000000)

# Build credentials from password and key
$Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminUser, ($ADPassword | ConvertTo-SecureString -Key $ADKey)

# Connect to Teams service
connect-microsoftteams -credential $Creds

# Add users to Team
Foreach ($User in $UsersForTeam) {

    Try {

    Add-TeamUser -GroupId 000000000000-0000000000000 -User $User -ErrorAction Stop
    write-host "Added: $User" -ForegroundColor Green
    Write-Log -Entry "Added: $User"

    }

    Catch {
    write-host "Existing: $User" -ForegroundColor Yellow

    }

}


#-----------------------------------------------------------[Cleanup]------------------------------------------------------------

$UserVariables = Get-Variable | Select-Object -ExpandProperty Name | Where-Object {$DefaultVariables -notcontains $_ -and $_ -ne "ExistingVariables"}
Remove-Variable $UserVariables
 

### Log End
Write-Log -Entry "Script ended ($ExitCode)."
### 

### Script END