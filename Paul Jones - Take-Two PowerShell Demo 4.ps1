#######################################################################################################################


#### PowerShell Script Template V1



#######################################################################################################################



# # Requires -Module %  >
# # Requires PowerShell Version 3 or above  >

<#

.SYNOPSIS
  

.DESCRIPTION


.INPUTS


.OUTPUTS


.NOTES
  Version:          %
  Author:           %
  Creation Date:    %
  Purpose/Change:   %
  Change Ref:       %
  Change Type:      %
  
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

  
#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script Version
$ScriptVersion = "%"
$ProjectName = "%"
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

###################################################################################################################################


#-----------------------------------------------------------[Cleanup]------------------------------------------------------------

$UserVariables = Get-Variable | Select-Object -ExpandProperty Name | Where-Object {$DefaultVariables -notcontains $_ -and $_ -ne "ExistingVariables"}
Remove-Variable $UserVariables
 

### Log End
Write-Log -Entry "Script DHCP ended ($ExitCode)."
### 

### Script END