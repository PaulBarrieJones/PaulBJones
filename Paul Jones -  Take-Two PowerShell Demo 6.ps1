#######################################################################################################################


#### Paul Jones - Take-Two - Script Example 6 - Exchange 2010 Service Desk Tool
####
#### This is an older script from 2016 and one of my first. The sctipt was written using PowerShell version 2. 
#### The tool allowed service desk to report to user on their mailbox usage and give them a small sie increase. 


#######################################################################################################################



## Mailbox Usage by Paul Jones v3


    ## Authntication and Remote Session into Exchange Server
          
    Try
    
    {
    
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://EXCHANGESERVER.MainDomain/powershell/" -Authentication Kerberos -Credential lsbu\Trevor1 -ErrorAction Stop

}
   
    Catch
        
    {
    Clear-Host

    Write-Host
    Write-Host "Password incorrect - logon failed" -ForegroundColor Red
    Write-Host 
    Pause
    
    Break


    }

    

    ### Once Authenticated Import Session
    
    Import-PSSession $Session -ErrorAction SilentlyContinue -WarningAction SilentlyContinue



    ### Main Menu
       
    Function MainMenu {  
    
    Clear-Host
    Write-Host 
    Write-Host "--------------------------------------------------------------"          -ForegroundColor Green
    Write-Host ""
    Write-Host "    User Mailbox Checker Main Menu"                                      -ForegroundColor Cyan
    Write-Host ""  
    Write-Host "    1. Username Finder"                                                  -ForegroundColor Green
    Write-Host "    2. Mailbox Usage Checker"                                            -ForegroundColor Green
    Write-Host "    0. Exit"                                                             -ForegroundColor Green
    Write-Host ""  
    Write-Host "--------------------------------------------------------------"          -ForegroundColor Green
    Write-Host ""
    $answer = Read-Host "Please Make a Selection"
                 


    if ($answer -eq 1){Get-UsernameInfo}  
    
    if ($answer -eq 2){Get-MailboxInfo} 

    if ($answer -eq 3){MainMenu}
    
    if ($answer -eq 0){Break}


 }


    ### Find Username Tool via Active Directory Services
    
    function Get-UsernameInfo {

    Import-Module ActiveDirectory

    Clear-Host
    
    Write-Host
    Write-Host "Staff Member Username Finder" -ForegroundColor Green
    Write-Host

    $username = read-host "Enter Surname of Staff Member (this may take some time)"
    
    get-aduser -filter *  -Properties name, samaccountname, extensionattribute5 | Where-Object {$_.name -like "*$username*"} | Where-Object {$_.extensionAttribute5 -eq "Staff"} | Select-Object Name, @{label=”Username”;expression={$_.SamAccountName}}

    Pause

    MainMenu

    }

    
    
    ### Check Staff Users Mailbox Status/Usage

    Function get-mailboxinfo {   

       
    Clear-Host
    
    Write-Host
    Write-Host "Exchange Mailbox Status" -ForegroundColor Green
    Write-Host
    

    $USER = Read-Host "Please enter username"

    Clear-Host

    Write-Host

    Try {
    
    $0 = invoke-command -session $Session {( get-mailbox $($args[0]) ).Name} -ArgumentList $user -ErrorAction Stop

    }

    Catch { Write-Host "Username incorrect" -ForegroundColor Red

    Pause

    Get-MailboxInfo
    
    }
    
    $1 = Invoke-Command -Session $session -ScriptBlock {(Get-Mailbox $($args[0]) ).ProhibitSendReceiveQuota.Value.ToMB()} -ArgumentList $user

    $2 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxStatistics $($args[0]) ).TotalItemSize.Value.ToMB()} -ArgumentList $user

    $3 = ( $2 / $1 * 100 ) 

    [INT]$4 = ( "{0:N0}" -f $3 )


    # $5 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxFolderStatistics -folderscope SentItems -Identity $($args[0]))} -ArgumentList $USER | Select-Object @{Name="SentTotal"; Expression = {$_.folderAndSubfolderSize}} -first 1

    $5 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxFolderStatistics -folderscope SentItems -Identity $($args[0]))} -ArgumentList $USER

    $6 = $5.folderAndSubfolderSize | select-object -First 1 | Out-String -Stream
        
    #[INT]$7 = $6 -replace ".B (.*)" , "" ###(REMOVE GB/MB too)

    $7 = $6 -replace "\(.*",""

    $8 = ( "{0:N0}" -f $7 )
    

    $9 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxFolderStatistics -folderscope DeletedItems -Identity $($args[0]))} -ArgumentList $USER

    $10 = $9.folderAndSubfolderSize | select-object -First 1 | Out-String -Stream
    
    #[INT]$7 = $6 -replace ".B (.*)" , "" ###(REMOVE GB/MB too)

    $11 = $10 -replace "\(.*",""

    $12 = ( "{0:N0}" -f $11 )
    
    
    Write-Host "Name: $0" -ForegroundColor Cyan
    Write-Host "Username: $User" -ForegroundColor Green
    Write-Host "Mailbox Quota: $1 MB" -ForegroundColor Green
    Write-Host "Mailbox Usage: $2 MB" -ForegroundColor Green
    Write-Host "Percentage used: $4%" -ForegroundColor Green
    Write-Host ""
    Write-Host "Further Information:" -ForegroundColor Cyan
    Write-Host "Sent Item Folder Size: $8" -ForegroundColor Green
    Write-Host "Deleted Items Folder size: $11" -ForegroundColor Green
    Write-Host ""
  
    Write-host "Current Mailbox Status:" -ForegroundColor Cyan
    If ( $4 -gt [INT]94 ) { Write-Host "User Send Blocked!" -ForegroundColor RED } else {Write-Host "Send Mail is Enabled" -ForegroundColor Green}
    If ( $4 -ge [INT]100 ) { Write-Host "USER RECEIVE BLOCKED!" -ForegroundColor WHITE -BackgroundColor RED }  else {Write-Host "Receive Email is Enabled" -ForegroundColor Green}
    write-host ""

    Write-Host "-----------------------------------"           -ForegroundColor Cyan
    write-host ""
    Write-Host "1. Boost this users quota by 100 MB"           -ForegroundColor Green
    Write-Host "2. Check another user account"                 -ForegroundColor Green
    write-host "3. Main menu"                                  -ForegroundColor Green
    write-host "4. Quit"                                       -ForegroundColor Gree
    Write-Host ""   
    Write-host "-----------------------------------"           -ForegroundColor Cyan
    Write-Host ""
    
    $answer2 = Read-Host "Please Make a Selection"
                 


    if ($answer2 -eq 1){Boost}  
    
    if ($answer2 -eq 2){get-mailboxinfo} 

    if ($answer2 -eq 3){MainMenu}
    
    if ($answer2 -eq 0){Break}

  
   
    
    Pause
    
    Clear-Host 

    #MainMenu


   }

  
  Function Boost {


    $IssueWarningQuota = Invoke-Command -Session $session -ScriptBlock {(Get-Mailbox $($args[0]) ).IssueWarningQuota.value.ToMB()} -ArgumentList $user
                
    $PRohibitSendQuota = Invoke-Command -Session $session -ScriptBlock {(Get-Mailbox $($args[0]) ).ProhibitSendQuota.value.ToMB()} -ArgumentList $user
        
    $ProhibitSendReceiveQuota = Invoke-Command -Session $session -ScriptBlock {(Get-Mailbox $($args[0]) ).ProhibitSendReceiveQuota.Value.ToMB()} -ArgumentList $user



    $IWQBOOST = $IssueWarningQuota + 100
    $PSQBOOST = $ProhibitSendQuota + 100
    $PSRBOOST = $ProhibitSendReceiveQuota + 100
        
    
    $IWQBOOSTMB = [STRING]$IWQBOOST+"MB"
    $PSQBOOSTMB = [STRING]$PSQBOOST+"MB"
    $PSRBOOSTMB = [STRING]$PSRBOOST+"MB"
        
    
    $collect = @()
        
    
    $collect += $user
    $collect += $IWQBOOSTMB
    $collect += $PSQBOOSTMB
    $collect += $PSRBOOSTMB

            
    Invoke-Command -Session $session -ScriptBlock {(Set-Mailbox -Identity $($args[0]) -IssueWarningQuota $($args[1]) -ProhibitSendQuota $($args[2]) -ProhibitSendReceiveQuota $($args[3]))} -ArgumentList $collect
  

    Clear-Host

    
    $0 = invoke-command -session $Session {( get-mailbox $($args[0]) ).Name} -ArgumentList $user -ErrorAction Stop
        
    $1 = Invoke-Command -Session $session -ScriptBlock {(Get-Mailbox $($args[0]) ).ProhibitSendReceiveQuota.Value.ToMB()} -ArgumentList $user

    $2 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxStatistics $($args[0]) ).TotalItemSize.Value.ToMB()} -ArgumentList $user

    $3 = ( $2 / $1 * 100 ) 

    [INT]$4 = ( "{0:N0}" -f $3 )


    # $5 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxFolderStatistics -folderscope SentItems -Identity $($args[0]))} -ArgumentList $USER | Select-Object @{Name="SentTotal"; Expression = {$_.folderAndSubfolderSize}} -first 1

    $5 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxFolderStatistics -folderscope SentItems -Identity $($args[0]))} -ArgumentList $USER

    $6 = $5.folderAndSubfolderSize | select-object -First 1 | Out-String -Stream
        
    #[INT]$7 = $6 -replace ".B (.*)" , "" ###(REMOVE GB/MB too)

    $7 = $6 -replace "\(.*",""

    $8 = ( "{0:N0}" -f $7 )
    

    $9 = Invoke-Command -Session $Session -ScriptBlock {(Get-MailboxFolderStatistics -folderscope DeletedItems -Identity $($args[0]))} -ArgumentList $USER

    $10 = $9.folderAndSubfolderSize | select-object -First 1 | Out-String -Stream
    
    #[INT]$7 = $6 -replace ".B (.*)" , "" ###(REMOVE GB/MB too)

    $11 = $10 -replace "\(.*",""

    #$12 = ( "{0:N0}" -f $11 )
    
    Write-Host
    Write-Host "[New Mailbox Status]" -ForegroundColor Green
    Write-Host
    Write-Host "Name: $0" -ForegroundColor Cyan
    Write-Host "Username: $User" -ForegroundColor Green
    Write-Host "Mailbox Quota: $1 MB" -ForegroundColor Green
    Write-Host "Mailbox Usage: $2 MB" -ForegroundColor Green
    Write-Host "Percentage used: $4%" -ForegroundColor Green
    Write-Host ""
    Write-Host "Further Information:" -ForegroundColor Cyan
    Write-Host "Sent Item Folder Size: $8" -ForegroundColor Green
    Write-Host "Deleted Items Folder size: $11" -ForegroundColor Green
    Write-Host ""
  
    Write-host "Current Mailbox Status:" -ForegroundColor Cyan
    If ( $4 -gt [INT]94 ) { Write-Host "User Send Blocked!" -ForegroundColor RED } else {Write-Host "Send Mail is Enabled" -ForegroundColor Green}
    If ( $4 -ge [INT]100 ) { Write-Host "USER RECEIVE BLOCKED!" -ForegroundColor WHITE -BackgroundColor RED }  else {Write-Host "Receive Email is Enabled" -ForegroundColor Green}
    write-host ""
    
    
    pause
   
    MainMenu


   }
    
    
    MainMenu


    #### End of Script