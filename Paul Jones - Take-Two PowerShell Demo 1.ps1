#######################################################################################################################


#### Paul Jones - Take-Two - Script example 1 - DHCP Configuration
####
#### This script was written to help progress a network upgrade project, the network team were entering all of the DHCP 
#### information manually into the DHCP servers from a CSV file. This work was duplicated for the two servers.
#### I wrote this script to automate the task, this saved a lot of time and ensured consistant naming conventions etc.
#### This script was needed urgently and took me around 2 hours to write, once it was tested on a test environment it 
#### ran successfully without errors. This script was written to be run once.
#### Please note, the script has been sanitised and any propriatary information has been renamed for security purposes.


#######################################################################################################################



# # Requires -Module DHCP.psm1  >
# # Requires PowerShell Version 3 or above  >

<#

.SYNOPSIS
  Add specifically formatted DHCP scopes to Microsoft DHCP servers from a CSV file.

.DESCRIPTION
  The network project team have published new network documentation into CSV file. Once imported into this
  script this data is calculated and formatted into scopes which are then created on the DHCP servers.
  Script to be run by infrastructure engineer, screen output can be observed during process.

.PARAMETER
  None

.INPUTS


.OUTPUTS


.NOTES
  Version:          2.0
  Author:           Paul Jones
  Creation Date:    14/10/2019
  Purpose/Change:   Project, planned work
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

    Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\DHCP\DHCP.psm1 -ErrorAction Stop

    }

        Catch {

        Clear-Host

        Write-Warning "DHCP.psm1 module not present on system, script cannot continue!"
        Write-Warning "To install modude (Server 2012 R2) run: (Add-WindowsFeature -Name DHCP -IncludeManagementTools)"

        BREAK

        }

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script Version
$ScriptVersion = "2.0"
$ProjectName = "DHCP"
$ProjectPath = "C:\PowerShell\Projects\$ProjectName"
$LogFilePath = "$ProjectPath\Logs\$ProjectName.log"
$ScriptPath = "$ProjectPath\Scripts\"

# DHCP Servers
$DHCPSVR1 = "SERVER1"
$DHCPSVR2 = "SERVER2"

# Log File Info

$LogFilePath = "$ProjectPath\Logs\$ProjectName.log"

# Payload file
$DHCPCSV = Import-Csv "$projectPath\Payload\Network.csv"


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

###################################################################################################################################

##################################
#### DATA Network Gather Start ###
##################################

$CustomArrayData = @()

foreach ($object in $DHCPCSV) {
    
    $ObjectAdd = New-Object -TypeName psobject
    
    # Variables for tidyness
    $postcode = $object.PostCode
    $Sitename = $object.SiteName
    
    $UIDSiteCode = $object.UIDSiteCode
    $ObjectAdd | Add-Member -MemberType NoteProperty -name Description -value "$UIDSiteCode - $sitename [$postcode]  |  Data (vlan2)"
    $ObjectAdd | Add-Member -MemberType NoteProperty -name PostCode -value $object.POSTCODE
    $objectAdd | Add-Member -MemberType NoteProperty -name UIDSiteCode -Value $object.UIDSiteCode
    $objectAdd | Add-Member -MemberType NoteProperty -name ScopeName -Value "Data (vlan 2)  |  $sitename ($UIDSiteCode)"
    $ObjectAdd | Add-Member -MemberType NoteProperty -name ScopeIPData -value $Object.ScopeIPData
    $ObjectAdd | Add-Member -MemberType NoteProperty -name StartRangeData -value $Object.StartRangeData
    $ObjectAdd | Add-Member -MemberType NoteProperty -name EndRangeData -value $object.EndRangeData
    $ObjectAdd | Add-Member -MemberType NoteProperty -name SubnetMask -value $object.SubnetMask
    
    $CustomArrayData += $ObjectAdd
    $2ndOctet = $object.StartRangeData.split(".")
    $2ndOctet = $2ndOctet[1]
   
 
### MAINSITE Exclusion Ranges for DATA (vlan2)

    # Get the exclusions for MAINSITE in a list
    $DHCPBuildExclusion = $object.MAINSITEDHCPExclusionsData.replace("x" ,$2ndOctet)
    # Split into an array
    $DHCPBuildExclusion1 = $DHCPBuildExclusion.split(":")
    
    # First of array split into range
    $DHCPMAINSITE1 = $DHCPBuildExclusion1[0]
    [ARRAY]$DHCPMAINSITE1 = $DHCPMAINSITE1.split("-")
    
    # FINAL for Script ::: MAINSITE DHCP Excl range 1
    $DHCPMAINSITE1Start = $DHCPMAINSITE1[0]
    $DHCPMAINSITE1End = $DHCPMAINSITE1[1]
    
    # Split into an array
    $DHCPBuildExclusion2 = $DHCPBuildExclusion.split(":")
    # Second of array split into range
    $DHCPMAINSITE2 = $DHCPBuildExclusion2[1]
    [ARRAY]$DHCPMAINSITE2 = $DHCPMAINSITE2.split("-")
    
    # FINAL for Script ::: MAINSITE DHCP Excl range 2
    $DHCPMAINSITE2Start = $DHCPMAINSITE2[0]
    $DHCPMAINSITE2End = $DHCPMAINSITE2[1]

    # Add 2 exclusion sets to Object
    $objectAdd | Add-Member -MemberType NoteProperty -name MAINSITEDHCPExclusions1Start -Value $DHCPMAINSITE1Start
    $objectAdd | Add-Member -MemberType NoteProperty -name MAINSITEDHCPExclusions1End -Value $DHCPMAINSITE1End
    $objectAdd | Add-Member -MemberType NoteProperty -name MAINSITEDHCPExclusions2Start -Value $DHCPMAINSITE2Start
    $objectAdd | Add-Member -MemberType NoteProperty -name MAINSITEDHCPExclusions2End -Value $DHCPMAINSITE2End
   
### SECONDSITE Exclusion Ranges for DATA (vlan3)

    # Get the exclusions for SECONDSITE in a list
    $DHCPBuildExclusion = $object.SECONDSITEDHCPExclusionsData.replace("x" ,$2ndOctet)
    # Split into an array
    $DHCPBuildExclusion1 = $DHCPBuildExclusion.split(":")
    
    # First of array split into range
    $DHCPSECONDSITE1 = $DHCPBuildExclusion1[0]
    [ARRAY]$DHCPSECONDSITE1 = $DHCPSECONDSITE1.split("-")
    
    # FINAL for Script ::: SECONDSITE DHCP Excl range 1
    $DHCPSECONDSITE1Start = $DHCPSECONDSITE1[0]
    $DHCPSECONDSITE1End = $DHCPSECONDSITE1[1]
    
    # Split into an array
    $DHCPBuildExclusion2 = $DHCPBuildExclusion.split(":")
    # Second of array split into range
    $DHCPSECONDSITE2 = $DHCPBuildExclusion2[1]
    [ARRAY]$DHCPSECONDSITE2 = $DHCPSECONDSITE2.split("-")
    
    # FINAL for Script ::: SECONDSITE DHCP Excl range 2
    $DHCPSECONDSITE2Start = $DHCPSECONDSITE2[0]
    $DHCPSECONDSITE2End = $DHCPSECONDSITE2[1]
    # Split into an array
    $DHCPBuildExclusion3 = $DHCPBuildExclusion.split(":")
    # Second of array split into range
    $DHCPSECONDSITE3 = $DHCPBuildExclusion3[2]
    [ARRAY]$DHCPSECONDSITE3 = $DHCPSECONDSITE3.split("-")
    
    # FINAL for Script ::: SECONDSITE DHCP Excl range 3
    $DHCPSECONDSITE3Start = $DHCPSECONDSITE3[0]
    $DHCPSECONDSITE3End = $DHCPSECONDSITE3[1]

    # Add 3 exclusion sets to Object
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions1Start -Value $DHCPSECONDSITE1Start
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions1End -Value $DHCPSECONDSITE1End
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions2Start -Value $DHCPSECONDSITE2Start
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions2End -Value $DHCPSECONDSITE2End
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions3Start -Value $DHCPSECONDSITE3Start
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions3End -Value $DHCPSECONDSITE3End

}

##################################
#### DATA Network Gather END #####
##################################

## =====================================

###################################
#### Voice Network Gather Start ###
###################################
$CustomArrayVoice = @()

foreach ($object in $DHCPCSV) {
    
    # Create new object to hold data
    $ObjectAdd = New-Object -TypeName psobject
    $Sitename = $object.SiteName
    $UIDSiteCode = $object.UIDSiteCode
    $postcode = $object.PostCode

    $ObjectAdd | Add-Member -MemberType NoteProperty -name Description -value "$UIDSiteCode - $sitename [$postcode]  |  Voice (vlan3)"
    $ObjectAdd | Add-Member -MemberType NoteProperty -name PostCode -value $object.POSTCODE
    $objectAdd | Add-Member -MemberType NoteProperty -name UIDSiteCode -Value $object.UIDSiteCode
    $objectAdd | Add-Member -MemberType NoteProperty -name ScopeName -Value "Voice (vlan 3)  |  $sitename ($UIDSiteCode)"
    $ObjectAdd | Add-Member -MemberType NoteProperty -name ScopeIPVoice -value $Object.ScopeIPVoice
    $ObjectAdd | Add-Member -MemberType NoteProperty -name StartRangeVoice -value $Object.StartRangeVoice
    $ObjectAdd | Add-Member -MemberType NoteProperty -name EndRangeVoice -value $object.EndRangeVoice
    $ObjectAdd | Add-Member -MemberType NoteProperty -name SubnetMask -value $object.SubnetMask

    $2ndOctet = $object.StartRangeVoice.split(".")
    $2ndOctet = $2ndOctet[1]

### MAINSITE Exclusion Ranges for Voice (vlan3)

    # Get the exclusions for MAINSITE in a list
    $DHCPBuildExclusion = $object.MAINSITEDHCPExclusionsVoice.replace("x" ,$2ndOctet)
    # Split into an array
    $DHCPBuildExclusion1 = $DHCPBuildExclusion.split(":")
    
    # First of array split into range
    $DHCPMAINSITE1 = $DHCPBuildExclusion1[0]
    [ARRAY]$DHCPMAINSITE1 = $DHCPMAINSITE1.split("-")
    
    # FINAL for Script ::: MAINSITE DHCP Excl range 1
    $DHCPMAINSITE1Start = $DHCPMAINSITE1[0]
    $DHCPMAINSITE1End = $DHCPMAINSITE1[1]
    # Add 1 exclusion set to Object
    $objectAdd | Add-Member -MemberType NoteProperty -name MAINSITEDHCPExclusions1Start -Value $DHCPMAINSITE1Start
    $objectAdd | Add-Member -MemberType NoteProperty -name MAINSITEDHCPExclusions1End -Value $DHCPMAINSITE1End

### SECONDSITE Exclusion Ranges for Voice (vlan3)

    # Get the exclusions for SECONDSITE in a list
    $DHCPBuildExclusion = $object.SECONDSITEDHCPExclusionsVoice.replace("x" ,$2ndOctet)
    # Split into an array
    $DHCPBuildExclusion1 = $DHCPBuildExclusion.split(":")
    
    # First of array split into range
    $DHCPSECONDSITE1 = $DHCPBuildExclusion1[0]
    [ARRAY]$DHCPSECONDSITE1 = $DHCPSECONDSITE1.split("-")
    
    # FINAL for Script ::: SECONDSITE DHCP Excl range 1
    $DHCPSECONDSITE1Start = $DHCPSECONDSITE1[0]
    $DHCPSECONDSITE1End = $DHCPSECONDSITE1[1]
    
 
    # Add 1 exclusion set to Object
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions1Start -Value $DHCPSECONDSITE1Start
    $objectAdd | Add-Member -MemberType NoteProperty -name SECONDSITEDHCPExclusions1End -Value $DHCPSECONDSITE1End
    $CustomArrayVoice += $ObjectAdd

}

###################################
#### Voice Network Gather End #####
###################################


##====================================================
### STAGE 2
##====================================================

###################################
#### Data Network Process Start ###
###################################

clear-host

# process first entry in source csv file
#$CustomArrayData = $CustomArrayData[0]

foreach ($object in $CustomArrayData) {

############## DHCP MAINSITE Data
  
    Try {
    
    ### Commands start
        Write-Host "adding Scope to:" $DHCPSVR1 -ForegroundColor Cyan
    
    New-DHCPScope -Server $DHCPSVR1 -Name $Object.ScopeName -Address $object.ScopeIPData -SubnetMask $object.SubnetMask -Description $object.Description -ErrorAction Stop
    
        Write-Host "SUCCESS: Scope added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding IP Range to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Add-DHCPIPRange -Server $DHCPSVR1 -scope $object.ScopeIPData -StartAddress $object.StartRangeData -EndAddress $object.EndRangeData -ErrorAction STOP
    
        Write-Host "SUCCESS: IP range added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding Exclusion range 1 to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $DHCPSVR1 -Scope $object.ScopeIPData -StartAddress $object.MAINSITEDHCPExclusions1Start -EndAddress $object.MAINSITEDHCPExclusions1End
    
        Write-Host "SUCCESS: Exclusion range 1 added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding Exclusion range 2 to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $DHCPSVR1 -Scope $object.ScopeIPData -StartAddress $object.MAINSITEDHCPExclusions2Start -EndAddress $object.MAINSITEDHCPExclusions2End
    
        Write-Host "SUCCESS: Exclusion range 2 added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding DHCP option 3 to:" $DHCPSVR1 -ForegroundColor Cyan

    Get-DHCPScope -Scope $object.ScopeIPData -Server $DHCPSVR1 | Set-DHCPOption -OptionID 3 -Value $object.EndRangeData -DataType IPADDRESS -ErrorAction Stop

        Write-Host "SUCCESS: Option 3 added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding voice DHCP option 44 to:" $DHCPSVR1 -ForegroundColor Cyan
        
    Get-DHCPScope -Scope $object.ScopeIPData -Server $DHCPSVR1 | set-DHCPOption -OptionID 44 -DataType IPADDRESS -Value 0.0.0.0
            
        Write-Host "SUCCESS: Option 44 cleared from" $DHCPSVR1 -ForegroundColor Green
        
    Get-DHCPScope -Scope $object.ScopeIPData -Server $DHCPSVR1 | Set-DHCPOption -OptionID 191 -Value "VLAN-A:3." -DataType String -ErrorAction Stop
    
        Write-Host "SUCCESS: Option 191 added to" $DHCPSVR1 -ForegroundColor Green
        write-host "SUCCESS: Added" $object.scopename "to" $DHCPSVR1 -ForegroundColor Green
    "------------------------"
        Write-log "SUCCESS: $DHCPSVR1 | $($object.ScopeIPData) | $($object).ScopeName"
    
### Commands end
    
    }
        Catch {
        write-host "FAILURE: Failed to add" $object.scopename "to" $DHCPSVR1 -ForegroundColor Red
        "------------------------"
        Write-Log  "FAIL: $DHCPSVR1 | $($object.ScopeIPData) | $($object.ScopeName)"
        
        }

############## DHCP SECONDSITE Data

    Try {
    
    ### Commands start
    Write-Host "adding Scope to:" $dhcpsvr2 -ForegroundColor Cyan
    
    New-DHCPScope -Server $dhcpsvr2 -Name $Object.ScopeName -Address $object.ScopeIPData -SubnetMask $object.SubnetMask -Description $object.Description #-ErrorAction Stop
    
        Write-Host "SUCCESS: Scope added to" $dhcpsvr2 -ForegroundColor Green
        Write-Host "adding IP Range to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Add-DHCPIPRange -Server $dhcpsvr2 -scope $object.ScopeIPData -StartAddress $object.StartRangeData -EndAddress $object.EndRangeData -ErrorAction STOP
    
        Write-Host "SUCCESS: IP range added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding Exclusion range 1 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $dhcpsvr2 -Scope $object.ScopeIPData -StartAddress $object.SECONDSITEDHCPExclusions1Start -EndAddress $object.SECONDSITEDHCPExclusions1End
    
        Write-Host "SUCCESS: Exclusion range 1 added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding Exclusion range 2 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $dhcpsvr2 -Scope $object.ScopeIPData -StartAddress $object.SECONDSITEDHCPExclusions2Start -EndAddress $object.SECONDSITEDHCPExclusions2End
    
        Write-Host "SUCCESS: Exclusion range 2 added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding Exclusion range 3 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $dhcpsvr2 -Scope $object.ScopeIPData -StartAddress $object.SECONDSITEDHCPExclusions3Start -EndAddress $object.SECONDSITEDHCPExclusions3End
    
        Write-Host "SUCCESS: Exclusion range 3 added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding DHCP option 3 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPData -Server $DHCPSVR2 | Set-DHCPOption -OptionID 3 -Value $object.EndRangeData -DataType IPADDRESS -ErrorAction Stop
    
        Write-Host "SUCCESS: Option 3 added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding voice DHCP option 44 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPData -Server $DHCPSVR2 | set-DHCPOption -OptionID 44 -DataType IPADDRESS -Value 0.0.0.0

        Write-Host "SUCCESS: Option 44 cleared on:" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding DHCP option 191 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPData -Server $DHCPSVR2 | Set-DHCPOption -OptionID 191 -Value "VLAN-A:3." -DataType String -ErrorAction Stop
    
        Write-Host "SUCCESS: Option 191 added to" $DHCPSVR2 -ForegroundColor Green
        write-host "SUCCESS: Added" $object.scopename "to" $DHCPSVR2 -ForegroundColor Green
        "------------------------"
        Write-Log "SUCCESS: $DHCPSVR2 | $($object.ScopeIPData) | $($object.ScopeName)"
    
    ### Commands end
    
    }
        Catch {
        
        write-host "FAILURE: Failed to add" $object.scopename "to" $DHCPSVR2 -ForegroundColor Red
        "------------------------"
        Write-log "FAIL: $DHCPSVR2 | $($object.ScopeIPData) | $($object.ScopeName)"
        
        }

################ SECONDSITE End
    
}

 
###################################
#### Data Network Process End   ###
###################################

 
## =====================================

 
#### VOICE START
##############################
#### Start Processing VOICE ###

clear-host

#$CustomArrayvoice = $CustomArrayvoice[1]

foreach ($object in $CustomArrayvoice) {

############## DHCP MAINSITE voice
    
    Try {
    
    ### Commands start
        Write-Host "adding voice voice Scope to:" $DHCPSVR1 -ForegroundColor Cyan
    
    New-DHCPScope -Server $DHCPSVR1 -Name $Object.ScopeName -Address $object.ScopeIPVoice -SubnetMask $object.SubnetMask -Description $object.Description -ErrorAction Stop
    
        Write-Host "SUCCESS: Scope added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding voice voice IP Range to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Add-DHCPIPRange -Server $DHCPSVR1 -scope $object.ScopeIPVoice -StartAddress $object.StartRangevoice -EndAddress $object.EndRangevoice -ErrorAction STOP
    
        Write-Host "SUCCESS: IP range added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding voice voice Exclusion range 1 to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $DHCPSVR1 -Scope $object.ScopeIPVoice -StartAddress $object.MAINSITEDHCPExclusions1Start -EndAddress $object.MAINSITEDHCPExclusions1End
    
        Write-Host "SUCCESS: Exclusion range 1 added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding voice DHCP option 3 to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR1 | Set-DHCPOption -OptionID 3 -Value $object.EndRangeVoice -DataType IPADDRESS -ErrorAction Stop
    
        Write-Host "SUCCESS: Option 3 added to" $DHCPSVR1 -ForegroundColor Green
        Write-Host "adding voice DHCP option 44 to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR1 | set-DHCPOption -OptionID 44 -DataType IPADDRESS -Value 0.0.0.0
    
        Write-Host "SUCCESS: Option 44 cleared on:" $DHCPSVR1 -ForegroundColor Green
        Write-Host "Adding voice DHCP option 191 to:" $DHCPSVR1 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR1 | Set-DHCPOption -OptionID 191 -Value "VLAN-A:3." -DataType String -ErrorAction Stop
  
        Write-Host "SUCCESS: Option 191 added to" $DHCPSVR2 -ForegroundColor Green
        write-host "SUCCESS: Added" $object.scopename "to" $DHCPSVR1 -ForegroundColor Green
    "------------------------"
        Write-log "SUCCESS: $DHCPSVR1 | $($object.ScopeIPVoice) | $($object.ScopeName)"
    
    ### Commands end
    
    }
        Catch {
        write-host "FAILURE: Failed to add" $object.scopename "to" $DHCPSVR1 -ForegroundColor Red
        "------------------------"
        Write-Log  "FAIL: $DHCPSVR1 | $($object.ScopeIPVoice) | $($object.ScopeName)"
        
        }
############# DHCP MAINSITE voice End
############################################################

############## DHCP SECONDSITE voice

    Try {

    ### Commands start
        Write-Host "adding voice Scope to:" $dhcpsvr2 -ForegroundColor Cyan
    
    New-DHCPScope -Server $dhcpsvr2 -Name $Object.ScopeName -Address $object.ScopeIPVoice -SubnetMask $object.SubnetMask -Description $object.Description #-ErrorAction Stop
    
        Write-Host "SUCCESS: Scope added to" $dhcpsvr2 -ForegroundColor Green
        Write-Host "adding voice IP Range to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Add-DHCPIPRange -Server $dhcpsvr2 -scope $object.ScopeIPVoice -StartAddress $object.StartRangevoice -EndAddress $object.EndRangevoice -ErrorAction STOP
    
        Write-Host "SUCCESS: IP range added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding voice Exclusion range 1 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Add-DHCPExclusionRange -Server $dhcpsvr2 -Scope $object.ScopeIPVoice -StartAddress $object.SECONDSITEDHCPExclusions1Start -EndAddress $object.SECONDSITEDHCPExclusions1End
    
        Write-Host "SUCCESS: Exclusion range 1 added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding voice DHCP option 3 to:" $DHCPSVR2 -ForegroundColor Cyan  
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR2 | Set-DHCPOption -OptionID 3 -Value $object.EndRangevoice -DataType IPADDRESS -ErrorAction Stop
    
        Write-Host "SUCCESS: Option 3 added to" $DHCPSVR2 -ForegroundColor Green
        
        Write-Host "adding voice DHCP option 44 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR2 | set-DHCPOption -OptionID 44 -DataType IPADDRESS -Value 0.0.0.0
            
        Write-Host "SUCCESS: Option 44 cleared on:" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding voice DHCP option 128 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR2 | Set-DHCPOption -OptionID 128 -Value "Nortel-i2004-A,10.10.33.2:4100,1,10;10.80.18.2:4100,1,10;10.80.0.16:44443,1,3." -DataType STRING -ErrorAction Stop

        Write-Host "SUCCESS: Option 128 added to" $DHCPSVR2 -ForegroundColor Green
        Write-Host "adding voice DHCP option 191 to:" $DHCPSVR2 -ForegroundColor Cyan
    
    Get-DHCPScope -Scope $object.ScopeIPVoice -Server $DHCPSVR2 | Set-DHCPOption -OptionID 191 -Value "VLAN-A:3." -DataType String -ErrorAction Stop
       
        Write-Host "SUCCESS: Option 191 added to" $DHCPSVR2 -ForegroundColor Green
        write-host "SUCCESS: Added" $object.scopename "to" $DHCPSVR2 -ForegroundColor Green
    "------------------------"
        Write-Log "SUCCESS: $DHCPSVR2 | $($object.ScopeIPVoice) | $($object.ScopeName)"
    
### Commands end
    
    }
        Catch {
       
        write-host "FAILURE: Failed to add" $object.scopename "to" $DHCPSVR2 -ForegroundColor Red
        "------------------------"
            "FAIL" + " : " + $DHCPSVR2 + " | " + $object.ScopeIPVoice + "  |  " + $object.ScopeName
        
        }

}

#### VOICE END

######################################################################################################

###

### Clear all user variables
$UserVariables = Get-Variable | Select-Object -ExpandProperty Name | Where-Object {$DefaultVariables -notcontains $_ -and $_ -ne "ExistingVariables"}
Remove-Variable $UserVariables
 

### Log End
Write-Log -Entry "Script DHCP ended ($ExitCode)."
### 

### Script END