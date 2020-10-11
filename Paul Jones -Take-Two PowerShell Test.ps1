
## PowerShell Scripting Evaluation 
## Completed by Paul Jones
## paulbjones@virginmedia.com
## Script tested on: PSVersion 5.1.19041.1

####### NOTE - Bug Alert! #########
# Question 8 - Custom array and CSV creation:
# The "BOS, Contractor" field is not populating with the correct figure of "1", I can not figure out why! All the others numbers are fine.
# I would investigate further but I have run out of time!
##################################

    Clear-host

    Write-host "Start" -ForegroundColor Cyan


#### Question 1:

    # Set data path
    $DataFilePath = "C:\Temp\Users.csv"

    # Collect user data in to variable
    $AllUsers = Import-Csv $DataFilePath

    Write-host 'Question 1: Users from CSV collected into array = $AllUsers' -ForegroundColor Green


### Question 2:

    # Get total number of users from array
    [INT]$AllUsersTotal = $Allusers.count

    # Display to screen
    Write-host "Question 2: Total users in array = $AllUsersTotal" -ForegroundColor Green


### Question 3:
    
    # Get total sum of all mailboxes from all users in array
    $AllUsersMailboxTotalGB = ($AllUsers.mailboxsizeGB | Measure-Object -Sum).sum

    Write-host "Question 3: Total mailbox size in GB = $AllUsersMailboxTotalGB" -ForegroundColor Green


### Question 4:

    # Compare all users email addresses and UPNs to detect differences
    $AllUsersUPNemailNotMatch = Compare-Object -referenceobject $($Allusers.userprincipalname) -DifferenceObject $($Allusers.EmailAddress) -CaseSensitive

    # Count the amount of users with differences
    $AllUsersUPNEmailNotMatchTotal = $AllUsersUPNEmailNotMatch.Inputobject.count

    Write-host "Question 4: Total users with UPN not matching EmailAddress = $AllUsersUPNEmailNotMatchTotal" -ForegroundColor Green


### Question 5:

    # Get users from the NYC site only
    $UsersNYC = $AllUsers | Where-Object {$_.site -eq "NYC"}
    
    # Get total sum of all mailboxes from all users in NYC site
    $NYCUsersMailboxTotalGB = ($UsersNYC.mailboxsizeGB | Measure-Object -Sum).sum

    Write-host "Question 5: Total mailbox size in GB at NYC site = $AllUsersUPNEmailNotMatchTotal" -ForegroundColor Green


### Question 6:

    # Get all users with account type of employees only
    $AllEmployees = $AllUsers | Where-Object {$_.AccountType -eq "Employee"} 

    # From list of employees get mailboxes greated than 10GB in size (it was necessary to convert number to integer here, PowerShell was confused)
    $AllEmployeesGT10 = $AllEmployees | Where-Object {[INT]$_.mailboxsizeGB -gt 10}

    Write-host 'Question 6: Users with mailbox of 10GB or more collected into array = $AllEmployeesGT10' -ForegroundColor Green


### Question 7:   
   
    # Get all users from NYC site with domain name @domain2.com
    $AllDomain2Users = $AllUsers | Where-Object {$_.Site -eq "NYC" -and $_.emailaddress -like "*domain2.com"}

    # Sort users by mailbox size
    $AllDomain2UsersSorted = $AllDomain2Users | Sort-Object -Descending mailboxsizegb

    # Get the username for each user from the emailaddress
    $Domain2Usernames = Foreach ($Domain2User in $AllDomain2UsersSorted) {

    ($Domain2User.emailaddress -split "@")[0]

    }

    # Join the results into a single line of string by joining on "space"
    $FinalListofUsernames = $Domain2Usernames -join " "

    Write-host "Question 7: String of usernames from @domain2.com =  $FinalListofUsernames " -ForegroundColor Green


### Question 8:  

    # Function to allow report to be run on demand from cmndlet
    Function Get-SiteReport {

    # Get the unique sites from the $AllUsers array
    $AllSites = $AllUsers | Select-Object site -Unique

    # Clear variable to prevent duplicated data
    $SiteInfo = $null

    # Create array
    $SiteInfo = @()
        
        Foreach ($Site in $AllSites) {

        ### Set site
        $Site = $Site.site

        ### Set user count per site
        $SiteUserCount = ($allusers | Where-Object {$_.site -eq $site}).count

        ### Set employee count per site
        $SiteEmployeeCount = ($allusers | Where-Object ({$_.site -eq $site -and $_.Accounttype -eq "Employee"})).count

        ### Set contractor count per site
        $SiteContractorCount = ($allusers | Where-Object ({$_.site -eq $site -and $_.Accounttype -eq  "Contractor"})).count

        ### Set mailbox total per site
        $SiteUsers = $AllUsers | Where-Object {$_.site -eq $site}
        $SiteUsersMailboxTotalGB = ($SiteUsers.mailboxsizeGB | Measure-Object -Sum).sum
        $SiteUsersMailboxTotalGBFormatted = "{0:N1}" -f $SiteUsersMailboxTotalGB

        ### Set mailbox average per site
        $SiteUsers = $AllUsers | Where-Object {$_.site -eq $site}
        $SiteUsersMailboxes = $siteusers.mailboxsizeGB 
        $SiteUserMailboxAverage = "{0:N1}" -f ($SiteUsersMailboxes | Measure-Object -Average).average

        # Create a custom object to hole each dataset
        $psObject = New-Object System.Object

        # Add values from previous commands to dataset
        $psObject | Add-Member -type NoteProperty -name Site -Value $Site
        $psObject | Add-Member -type NoteProperty -name TotalUserCount -Value $SiteUserCount
        $psObject | Add-Member -type NoteProperty -name EmployeeCount -Value $SiteEmployeeCount
        $psObject | Add-Member -type NoteProperty -name ContractorCount -Value $SiteContractorCount
        $psObject | Add-Member -type NoteProperty -name TotalMailboxSizeGB -Value $SiteUsersMailboxTotalGBFormatted 
        $psObject | Add-Member -type NoteProperty -name AverageMailboxSizeGB -Value $SiteUserMailboxAverage

        # Add dataset to array
        $SiteInfo += $psObject

        }

        # Output the array as a CSV file with date it was generated
        
        $date = Get-Date -Format ddMMyy

        $siteInfo | Export-csv "C:\Temp\SiteReport-$date.csv" -NoTypeInformation

}

    # Run the function to generate the CSV
    Get-SiteReport

    Write-host "Question 8: Site report output to C:\Temp\SiteReport.csv" -ForegroundColor Green

    Write-host "END" -ForegroundColor Cyan




