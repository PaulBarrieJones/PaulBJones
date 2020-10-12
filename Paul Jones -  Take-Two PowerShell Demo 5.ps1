#######################################################################################################################


#### Paul Jones - Take-Two - Script Example 5 - WSUS Automation
####
#### This is an older script from 2017. It collates all servers in the live environment (connected to domain) and 
#### prepares the systems for the weekly patch cycle. It connects and configures VSSphere tags, AD, Group Policy WSUS and 
#### produces a report to an Excel template with the inteneded patch cycle, and the updates to be installed. This 
#### script was run against around 1000 VMs and physical boxes, a seperate script was used to control the updates.


#######################################################################################################################


  # This script will gather all Windows servers on the live MainDomain network and check which Security and Critical updates
  # are required. It will then produce a report to be emailed to the application owners.
  # Finally it will flush and repopulate the OUs, Group Policy and WSUS servers based upon the Patch Cycle metadata.
  # When ready the relevent Group policies will need to be manually engaged by the administrator.
  #
  # By Paul Jones
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  # Prepare logging variables and remove old logs
  
  Set-ExecutionPolicy Unrestricted -Force
  
  [string]$DateYMD = "" + (Get-Date).Year + "-" + (Get-Date).Month + "-" + (Get-Date).Day

  Remove-Item D:\SCRIPTS\LOGGING\*
  Remove-Item D:\SCRIPTS\OUTPUT\*
  
  $log =  "D:\scripts\LOGGING\$dateYMD-log.txt"
  $ServersWithErrorsLog = "D:\scripts\LOGGING\ServersWithErrorsLog.txt"

  # Exceptions - add server name here to remove it from processing
  $ex1 = "gemini"
  $ex2 = $null
  $ex3 = $null
  $ex4 = $null
  $ex5 = $null
  


  #===========================================
  # Connect to WSUS server and load assemblies
  #===========================================
      
  $wsusserver = 'WSUS2012'
  
  #Load required assemblies            
  
  [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
  
  $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$False)


  


# =================
# STAGE 1 - VSphere
# =================
# =================

# =====================================
# Connect to PowerCLI (VSphere) service   
# =====================================

    
  # Clear-Host

  Remove-Item D:\SCRIPTS\OUTPUT\* -Force

  Write-Host
  
  Write-Host "Logging to the VSphere system (Connecting...):" -ForegroundColor Green

  
  # Add VMWare PSSnapin unless it is already loaded
  
  $CheckCLi = Get-PSSnapin | Where-Object {$_.Name -like "VMware.VimAutomation.Core"}
          
      if ($null -eq $CheckCLi) {Add-PSSnapin VMware.VimAutomation.Core}
      
        
  # Logon to VCenter server - [THIS WILL USE AUTO LOGON]

  # ALREADY DONE  - Read-host -assecurestring | convertfrom-securestring | out-file "C:\WSUS\Secure\autoadminpwd.txt"
        
  $username = "adminAccount"
      
  $password = Get-Content "C:\WSUS\Secure\adminAccount.txt" | ConvertTo-Securestring
      
  $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password

  Connect-VIserver -Server VCenterServer -ErrorAction Stop -WarningAction SilentlyContinue -Credential $cred

        

  # ========================================================
  # Gather viable Windows OS servers from VSphere and Filter
  # ========================================================

  $ServerList = Get-VM -Location  "Awaiting Comissioning" , "Dev/Test VMs" , "Production VMs"  | 
      
          Where-Object {$_.Guest.OSFullName -like "*Microsoft Windows Server*"} | 
      
          Where-Object {$_.PowerState -ne "PoweredOff"} | 
  
          Where-Object {$_.NetworkAdapters.NetworkName -notlike "*Privnet*"} | 
  
          Where-Object {$_.NetworkAdapters.NetworkName -notlike "*TestNet*"} |
  
          Where-Object {$_.NetworkAdapters.NetworkName -notlike "*Staging*"} | 
      
          Where-Object {$_.Guest.Hostname -notlike "*.MainDomaintest.net*"} | 
          
          where-object {$_.name -notlike $ex1} -WarningAction Ignore
          
          
# Log Processed Servers
#################################################################
$ServerList.name | sort | write-host -ForegroundColor Green
Write-Output "-------------------------------------------------"  >> $log  
Write-Output "Date : $DateYMD"                                    >> $log
Write-Output "-------------------------------------------------"  >> $log  
Write-Output "Processed Servers"                                  >> $log          
Write-Output "-------------------------------------------------"  >> $log
$ServerList.name                                          >> $log
Write-Output ""                                                   >> $log
#################################################################



  # ===============================
  # Load and prepare Excel Template 
  # ===============================

  $erroractionpreference = "Continue"
  
  # Load Excel application into memory
  
  $Excel = New-Object -comobject Excel.Application
  
  # Show spreadsheet on screen - hidden or visable - has not impact either way
  
  $Excel.visible = $True 

  # Open exisiting Excel template
  
  $workbook = $excel.Workbooks.Open("D:\SCRIPTS\Templates\WSUSExcelApplicationOwnersUpdatesRequiredTemplatev5.xltx")
  
  # Select the correct WorkSheet: 0 = Servers and Owners | 1 = Servers and Updates Required
  
  $worksheet = $workbook.Worksheets.item(1)
  
  # Switch to different sheet on the screen (makes no difference to the script)
  
  # Set starting Excel row (from the top)
  
  $cell = 8


  # ==================================================================================================================
  # Prepare Data in $Serverlist for output to WSUS server, communications document and Group Policy scope settings
  # ==================================================================================================================

  
  $VMArray = @()

  
  # Pull info from each server in $serverlist

  foreach ($servername in $Serverlist)
  
      {
  
      # =====================================
      # Prepare server details into variables
      # =====================================

      # Get Application Owner from the array of custom data
  
      $AppownerPrepare = $servername.customfields -match "Application Owner"
  
      $AppownerReady = $AppownerPrepare.value 
  
      # Seperate server name and description into two fields
      
      $Description = $Servername.Name -split " - " 
      
      #$Description = $Servername.Name.Split(" - ")

      # Get data from custome field "Patch Cycle"
      
      $PatchCycle = $servername.CustomFields -Match "Patch Cycle"

      # Get data from custome field "Patch Mode"

      $PatchMode = $servername.CustomFields -Match "Patch Mode"


      # =============================================================================
      # Build custom output array from variables and poplulate spreadsheet on-the-fly
      # =============================================================================

      $Object = New-Object PSObject
  
          
          $Object | Add-Member "ServerName" $description[0].Substring(1)

                  [STRING]$excel.cells.item($cell,2) = $Description[0].Substring(1) ############ Output to EXCEL >>>
  
          $Object | Add-Member "Description" $description[1].TrimStart()
  
                  [STRING]$excel.cells.item($cell,3) = $Description[1].TrimStart() ############ Output to EXCEL >>>
  
          $Object | Add-Member "ApplicationOwner" $AppownerPrepare.Value
  
                  [STRING]$excel.cells.item($cell,4) = $AppownerPrepare.Value ############ Output to EXCEL >>>
      
          $Object | Add-Member "HostnameDNS" $Servername.Guest.HostName #.Replace(".MainDomain.com","").ToUpper()
  
                  [STRING]$excel.cells.item($cell,5) = $Servername.Guest.HostName.Replace(".MainDomain.com","").ToUpper() ############ Output to EXCEL >>>
  
          $object | Add-Member "PatchCycle" $PatchCycle.Value

                  [STRING]$excel.cells.item($cell,6) = $PatchCycle.Value ############ Output to EXCEL >>>

          $object | Add-Member "PatchMode" $PatchMode.Value

                  [STRING]$excel.cells.item($cell,7) = $PatchMode.Value ############ Output to EXCEL >>>
  
      
      $VMArray += $Object

      $cell ++

      }


# =====================
# STAGE 2 - WSUS Server
# =====================
# =====================



  #===========================================
  # Connect to WSUS server and load assemblies
  #===========================================
      
  $wsusserver = 'WSUSServer'
  
  #Load required assemblies            
  
  [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
  
  $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$False)


  
  #=======================================================
  # Clear existing VMs from Groups in WSUS to refresh list
  #=======================================================
  #################################################################
  Write-Output ""                                                   >> $log
  Write-Output "-------------------------------------------------"  >> $log  
  Write-Output "Clear Existing Servers from WSUS Server:"           >> $log          
  Write-Output "-------------------------------------------------"  >> $log
  #################################################################  
  
  
  # Get processing groups

  $PatchCycleGroups = $wsus.GetComputerTargetGroups() | Where-object {($_.Name -eq "PatchCycle1") -or ($_.name -eq "PatchCycle2") -or ($_.name -eq "PatchCycle3")}
  
  # Gather existing PCs from group

  $GetPCList = $PatchCycleGroups.GetComputerTargets()
  
  # Remove all existing PCs and rebuild
        
      
      ForEach ($client in $GetPCList) 
  
          {
          
          $client.FullDomainName | write-host -NoNewline ; write-host " - Clearing from WSUS Server" -ForegroundColor Red

          $PatchCycleGroups.RemoveComputerTarget($client)

          #$client.FullDomainName | Write-Output
          
          }


  #==================================================================
  # Add the captured servers into their correct groups by Patch Cycle
  #==================================================================

  
  # Define the WSUS groups into variables
  
  $group1 = $wsus.GetComputerTargetGroups() | Where-Object {$_.Name -eq "Patchcycle1"}

  $group2 = $wsus.GetComputerTargetGroups() | Where-Object {$_.Name -eq "Patchcycle2"}
              
  $group3 = $wsus.GetComputerTargetGroups() | Where-Object {$_.Name -eq "Patchcycle3"}
  
  
  Clear-host

# Log Servers Patch Cycle
#################################################################    
Write-Output "-------------------------------------------------"  >> $log  
Write-Output "Server Patch Cycles"                                >> $log          
Write-Output "-------------------------------------------------"  >> $log
#################################################################

  clear-host
  
  Foreach ($computer in $VMarray) 

  { 
      
      if ($Computer.'PatchCycle' -eq "1") 
  
          { 
      
          write-host "Adding to Patch Cycle 1 - " -ForegroundColor Green -NoNewline ; $computer.ServerName
          write-output $computer."ServerName" "Adding to Patch Cycle 1" >> $log
                      
              try { $client1 = $wsus.GetComputerTargetByName($computer."hostnameDNS") ; $group1.AddComputerTarget($client1)  } 
          
                  catch { 
                                      
                          write-host "DNS Issue? - " -ForegroundColor Red -NoNewline ; $computer.ServerName
                          Write-Output "===========" >> $log
                          write-output "DNS Issue?" >> $log 
                          Write-Output "===========" >> $log
                          Write-Output $computer.Servername >> $ServersWithErrorsLog
              
                        }
      
          }  

              elseif ($computer.'PatchCycle' -eq "2") 
      
              { 
          
              write-host "Adding to Patch Cycle 2 - " -ForegroundColor Green -NoNewline ; $computer.ServerName
              write-output $computer."servername" "Adding to Patch Cycle 2" >> $log

                  try { $client2 = $wsus.GetComputerTargetByName($computer."hostnameDNS") ; $group2.AddComputerTarget($client2) } 
              
                      catch { 

                              write-host "DNS Issue? - " -ForegroundColor Red -NoNewline ; $computer.ServerName
                              Write-Output "===========" >> $log
                              write-output "DNS Issue?" >> $log 
                              Write-Output "===========" >> $log
                              Write-Output $computer.Servername >> $ServersWithErrorsLog
              
                            }
              

              }
              
                    elseif ($computer.'PatchCycle' -eq "3") 
      
                  { 
          
                  write-host "Adding to Patch Cycle 3 - " -ForegroundColor Green -NoNewline ; $computer.ServerName
                  write-output $computer."servername" "Adding to Patch Cycle 3" >> $log

                      try { $client3 = $wsus.GetComputerTargetByName($computer."hostnameDNS") ; $group3.AddComputerTarget($client3) } 
                  
                          catch {
                      
                                  write-host "DNS Issue? - " -ForegroundColor Red -NoNewline ; $computer.ServerName
                                  Write-Output "===========" >> $log
                                  write-output "DNS Issue?" >> $log 
                                  Write-Output "===========" >> $log 
                                  Write-Output $computer.Servername >> $ServersWithErrorsLog
                              
                                  }
          
                  }
                  
                  
                      else 

                      { 
              
                      write-output $computer."servername" "Patch Cycle 0 or error?" -ForegroundColor green >> $log
                      # Try Fixing WSUS Client
                                                                                                                  
                      }

  }




# ==================================
# STAGE 3 - Group Policy Permissions
# ==================================
# ==================================


#Function WSUSGroupPolicy-Flush {
  
  
  # Clear all existing computers from the Group Policy Objects
      
  # Logging
      
  #################################################################
  Write-Output ""                                                   >> $log
  Write-Output "-------------------------------------------------"  >> $log  
  Write-Output "WSUS PCs Permissions Flush:"                        >> $log          
  Write-Output "-------------------------------------------------"  >> $log
  #################################################################
  
  
  #####################################################################################################
  # Patch Cycle 1
  
  $GPOPatchCycle1 = Get-GPPermissions -Name "WSUS2012 Patch Cycle 1" -all | where {$_.trustee.name -like "*$*"}

  Write-Output "From Patch Cycle 1" >> $log
  Write-Output "--------------------------------------------------" >> $log
  
      ForEach ($Permission in $GPOPatchCycle1)
  
      { 
            
          Set-GPPermissions -Name "WSUS2012 Patch Cycle 1" -TargetName $permission.Trustee.Name.Replace("$","") -TargetType Computer -PermissionLevel None 
      
          #TEST write-host "Clearing from GP - " -ForegroundColor Red -NoNewline ; $permission.Trustee.Name.Replace("$","")
          
          Write-Output $Permission.Trustee.Name.Replace("$","") >> $log
                      
              
      }

  ########################################################################################################
  ########################################################################################################
  # Patch Cycle 2

  $GPOPatchCycle2 = Get-GPPermissions -Name "WSUS2012 Patch Cycle 2" -All | where {$_.trustee.name -like "*$*"}

  Write-Output "--------------------------------------------------" >> $log
  Write-Output "From Patch Cycle 2"                                 >> $log
  Write-Output "--------------------------------------------------" >> $log
  
      ForEach ($Permission in $GPOPatchCycle2)
  
      {                 
                  
          Set-GPPermissions -Name "WSUS2012 Patch Cycle 2" -TargetName $permission.Trustee.Name.Replace("$","") -TargetType Computer -PermissionLevel None 
      
          Write-Output $permission.trustee.name.replace("$","") >> $log
      
          
      }

  #################################################################################################
  ########################################################################################################
  # Patch Cycle 3 

  $GPOPatchCycle3 = Get-GPPermissions -Name "WSUS2012 Patch Cycle 3" -All | where {$_.trustee.name -like "*$*"}
    
  Write-Output "--------------------------------------------------" >> $log
  Write-Output "From Patch Cycle 3"                                 >> $log
  Write-Output "--------------------------------------------------" >> $log
      
      ForEach ($Permission in $GPOPatchCycle3)
  
      { 
      
          Set-GPPermissions -Name "WSUS2012 Patch Cycle 3" -TargetName $permission.Trustee.Name.Replace("$","") -TargetType Computer -PermissionLevel None 
      
          Write-Output $permission.trustee.name.replace("$","") >> $log
      
      }

  #################################################################################################
  # END OF FUNCTION

  
    
  
  # Add computers to relevent group policy

      # Logging
      
      Write-Output ""                                                   >> $log
      Write-Output "-------------------------------------------------"  >> $log  
      Write-Output "WSUS Group Policy - Computer added:"                >> $log          
      Write-Output "-------------------------------------------------"  >> $log

  
  ForEach ($Computer in $VMArray)

  {
        
        If ($Computer.'PatchCycle' -eq "1") 
        
            {
          
            Set-GPPermissions -Name "WSUS2012 Patch Cycle 1" -TargetName $computer.HostnameDNS.replace(".MainDomain.com","") -TargetType Computer -PermissionLevel GpoApply

            #Write-output "Added to GP 1" >> $log 
            
            Write-host $computer.HostnameDNS.replace(".MainDomain.com","") "- Added to: WSUS2012 Patch Cycle 1" -ForegroundColor Cyan
            Write-Output $computer.HostnameDNS.replace(".MainDomain.com","") "- Added to: WSUS2012 Patch Cycle 1" >> $log

            }

          
              Elseif ($Computer.'PatchCycle' -eq "2")
            
                  {Set-GPPermissions -Name "WSUS2012 Patch Cycle 2" -TargetName $computer.HostnameDNS.replace(".MainDomain.com","") -TargetType Computer -PermissionLevel GpoApply

                    #Write-Output "$computer Added to GP 2" >> $log 
                    
                    Write-host $computer.HostnameDNS.replace(".MainDomain.com","") "- Added to: WSUS2012 Patch Cycle 2" -ForegroundColor Cyan
                    Write-Output $computer.HostnameDNS.replace(".MainDomain.com","") "- Added to: WSUS2012 Patch Cycle 2" >> $log

                    }

                      
                          Elseif ($Computer.'PatchCycle' -eq "3")

                              {Set-GPPermissions -Name "WSUS2012 Patch Cycle 3" -TargetName $computer.HostnameDNS.replace(".MainDomain.com","") -TargetType Computer -PermissionLevel GpoApply

                                #Write-Output "$computer Added to GP 3" >> $log 
                                
                                Write-host $computer.HostnameDNS.replace(".MainDomain.com","") "- Added to: WSUS2012 Patch Cycle 3" -ForegroundColor Cyan
                                Write-Output $computer.HostnameDNS.replace(".MainDomain.com","") "- Added to: WSUS2012 Patch Cycle 3" >> $log
                                
                                }

                                  Else {}


  }



  
  
  
  # Gather information on all updates into variable
  
  #$updates = $wsus.GetUpdates()

  $updates = $wsus.GetUpdates() | Where-Object {$_.isapproved -eq "true"} | Where-Object {$_.UpdateClassificationTitle -eq "Critical Updates" -or $_.UpdateClassificationTitle -eq "Security Updates"}
  
  # Create a WSUS computer scope to gather all computers into a variable
  
  $computerscope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope

  # Add computers to the scope
  
  #$wsus.GetComputerTargets($computerscope)



  #####
  
  # Run report on computers and updates required filtering out irrelevent updates
  clear-host

  write-host "Gathering data - this can take a long time" -ForegroundColor Red

  $gather = $updates.GetUpdateInstallationInfoPerComputerTarget($ComputerScope) | 

      Where-Object {$_.updateinstallationstate -ne "NotApplicable"} | 

      Where-Object {$_.updateinstallationstate -ne "Installed"} | 

      Where-Object {$_.updateapprovalaction -ne "NotApproved"} | 

      Select-Object @{L=’Client';E={$wsus.GetComputerTarget(([guid]$_.ComputerTargetId)).FulldomainName}}, 
      
      @{L=’Update';E={$wsus.GetUpdate(([guid]$_.UpdateId)).knowledgebasearticles}} | Sort-Object client



      #=========
      # EXCEL 2
      #=========
      
      # Switch Excel spreadsheet to Sheet 2
      
      $worksheet = $workbook.Worksheets.item(2)
                
      # Visually shift to sheet 2 (debug)
      
      $worksheet.select(2)

      # Start from cell 8
      
      $cell2 = 8 
            
      # Create second array
      
      $VMArray2 = @()
    
  
          # Use Array of objects to export to excel
          
          ForEach ($Instance in $gather) 
          
                {
  
  
                  $Object2 = New-Object PSObject
  
          
                  $Object2 | Add-Member "Server Name" $Instance.client.Replace(".MainDomain.com","")

                          [STRING]$excel.cells.item($cell2,2) = $Instance.Client.Replace(".MainDomain.com","") ############ To EXCEL
  
                  $Object2 | Add-Member "KB Reference" $Instance.Update
  
                          [STRING]$excel.cells.item($cell2,3) = $Instance.Update ############ To EXCEL
  
          
                  # Add objects to array
                  
                  $VMArray2 += $Object2

                  # Add a number to choose line in excel (started at 8)
                  
                  $cell2 ++


                }


    # =======================
    # Finalise and write file
    # =======================


    # Set filename and paths

      # Set filename for output using date
  
      $filename = "SecurityPatchesRequired{0:yyyMMdd-HHmm}" -f (Get-Date)
      
      # Set ouput path
  
      $OutputPath = "D:\SCRIPTS\OUTPUT\"
                      
      # Save the file
  
      $workbook.SaveAs($OutputPath+$Filename)
  
      # Quit Excel
  
      $Excel.Quit() 

      
      Clear-host
      
      Write-Host
      Write-Host "The following list of server have WSUS connection issues, please investigate and repair:" -ForegroundColor Cyan
      Write-Host
      $ServersWithIssues = Get-Content "D:\scripts\LOGGING\ServersWithErrorsLog.txt"
      $ServersWithIssues | Write-Host -ForegroundColor Red
      
      
      # List for reboot script

      # Remove old list
      
      Remove-Item D:\SCRIPTS\OUTPUT\RebootScript-PatchCycle*.* -Force

      # Create new lists

      ($VMArray | Where-Object {$_.patchcycle -eq "1"}).hostnamedns | out-file D:\SCRIPTS\OUTPUT\RebootScript-PatchCycle1.txt
      ($VMArray | Where-Object {$_.patchcycle -eq "2"}).hostnamedns | out-file D:\SCRIPTS\OUTPUT\RebootScript-PatchCycle2.txt
      ($VMArray | Where-Object {$_.patchcycle -eq "3"}).hostnamedns | out-file D:\SCRIPTS\OUTPUT\RebootScript-PatchCycle3.txt
      ($VMArray | Where-Object {$_.patchcycle -eq "0"}).hostnamedns | out-file D:\SCRIPTS\OUTPUT\RebootScript-PatchCycle0.txt
      
      
      
      Write-Host 
      Write-Host "Work Complete - Excel Spreadsheet is ready." -ForegroundColor Green
      Write-Host "Please go to: $outputpath" -ForegroundColor Cyan
      Write-Host
      
      Write-Output "----------" >> $log
      Write-Output "End of Log" >> $log
      Write-Output "----------" >> $log

      
  #==============
  # END of script
  #==============