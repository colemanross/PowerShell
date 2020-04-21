function Login-ToOffice365 {

    # If there is currently no Office 365 PS Session open
    if (-Not(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })){
        # Attempt to connect to Office 365
        try {
            # Get logged in username and append "ed.ac.uk
            $userName = $env:USERNAME + "@ed.ac.uk"
            # Get user credential
            $UserCredential = Get-Credential -UserName $userName -Message "Please enter your Office 365 password"
            # Create new PS Session for Office 365
            $Office365SessionLo = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
            # Start the session
            Import-PSSession $Office365Session -ErrorAction Stop -DisableNameChecking > $null
            # Display message for user
            Write-Host "`nSuccessfully connected to Office 365. Please wait a moment for prompt to appear...`n" -ForegroundColor Green
            # Wait 2 seconds
            Start-Sleep -s 2
        }
        # Unable to login - most likely incorrect password or permissions
        catch {        
            Write-Host "`nCannot connect to Office 365. Have you entered the correct username and password? Do you have the correct permissions?`n" -ForegroundColor Red
            Write-Host "Cause: $_." -ForegroundColor Red
            "`n"
            return
        }
    }
    # Else Office 365 session is already open
    else {
        Write-Host "You are already connected to Office 365." -ForegroundColor Yellow
    }
}

function Logout-OfOffice365 {
    
    # Internal function to see if currently logged into Office 365
    function Logout() {
        # Check to see if there is currently an Office 365 session open
        if (-Not(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })){
            # If there is no current session then display message
            Write-Host "`nYou are not currently logged in to Office 365`n" -ForegroundColor Yellow
        }
        # Else there is a session open
        else {
            # Display user message
            Write-Host "`nAttempting to logout of Office 365.......`n"
            # Get the session
            $Office365Session = Get-PSSession
            # Logout of the session
            Remove-PSSession $Office365Session
            # Display user message
            Write-Host "`nLogout Successful!`n" -ForegroundColor Green    
        }    
    }

    # Begin main loop
    do {
        # Prompt user
        $loggedIn = Read-Host "`nDo you want to logout of Office 365 (y/n) ? "
        # Convert user entry to upper case       
        $loggedIn = $loggedIn.ToUpper()
        # Switch through possible replies
        switch ($loggedIn) {
        # If answer is yes
            "Y"{
                # Logout of office 365
                Logout
                # Display message
                Write-Host "`nQuitting script....`n"
                # Exit script
                return                           
            }
            # If answer is no
            "N"{
                # Display message
                Write-Host "`nQuitting script but staying logged into Office 365....`n"
                # Quit script
                return
            }
            # If reply is invalid
            default {Write-Host "`n*** ERROR *** Please enter y or n " -ForegroundColor Red}
        }
    # Loop until entry is equal to Y or N
    } until ($loggedIn -eq "Y" -or $loggedIn -eq "N")
}

function Display-MailboxInfo {
    
    # Make sure we are logged in
    Login-ToOffice365

    # Begin main program loop. Loop until valid uun is supplied
    do {
        
        $uun = Read-Host "Please enter username of account (or CTRL+C to quit)"
      
        $userExists = (Check-Mailbox $uun)
        # If account does not exist then display error message 
        if ($userExists -eq "error") {
            # Display error message
            Write-Host "`nCannot find mailbox for username $uun`n" -ForegroundColor Red
        }
    } until($userExists -ne "error")
    
    Write-Host "`n"
    # Display header
    "Mailbox Info for $uun"
    "===================================="
    # Create 2 arrays so that we can customise the output,
    # one to hold the output for the command 'Get-MailboxStatistics', and one for 'Get-Mailbox'.
    $a = @{Expression={$_.DisplayName};Label="Display Name                 "},@{Expression={$_.TotalItemSize};Label="Account Usage"},@{Expression={$_.TotalDeletedItemSize};Label="Deleted Items Size"},@{Expression={$_.LastLogonTime};Label="Last logged on"}
    $b = @{Expression={$_.PrimarySMTPAddress};Label="Primary SMTP Address"},@{Expression={$_.DeliverToMailBoxAndForward};Label="Keep Forward Message In Inbox"},@{Expression={$_.ForwardingSmtpAddress};Label="Forwarding Address"},@{Expression={$_.IsMailBoxEnabled};Label="Is Mailbox Enabled"}
    
    # Use command 'Get-MailboxStatistics' and then 'Get-Mailbox', supplying username. Format as a list using array, and then cut any empty output lines at the end.
    (Get-MailboxStatistics $uun | Format-List $a | Out-String).Trim()
    (Get-Mailbox $uun | Format-List $b | Out-String).Trim()
    
    # Call 'Logout-OfExchange' to logout of Office 365
    Logout-OfOffice365
}

function Display-MailboxUsage {
    
    # Make sure we are logged in to Office365
    Login-ToOffice365

    # Begin main program loop. Loop until valid uun is supplied
    do {        
        $uun = Read-Host "Please enter username of account (or CTRL+C to quit)"
        # Convert username to upper case
        $uun = $userName.ToLower()
        $userExists = (Check-Mailbox $uun) 
    } until($userExists -eq "User exists")     

    # Display Usage
    Get-MailboxStatistics $uun | Format-List DisplayName,TotalItemSize,TotalDeletedItemSize,ItemCount,DeletedItemCount
    # Logout of Office 365
    Logout-OfOffice365                
}

function Display-PrimarySMTPAddress {
    # Make sure we are logged in to Office365
    Login-ToOffice365

    # Begin main program loop. Loop until valid uun is supplied
    do {        
        $uun = Read-Host "Please enter username of account (or CTRL+C to quit)"
        # Convert username to upper case
        $uun = $userName.ToLower()
        $userExists = (Check-Mailbox $uun) 
    } until($userExists -eq "User exists")

    # Create array to hold output
    $output = @{Expression={$_.DisplayName};Label="Display Name"},@{Expression={$_.PrimarySmtpAddress};Label="Primary SMTP Address"}
    try {
        # Display header
        Write-Host "Primary SMTP Address for $uun"
        Write-Host "===================================="
        # Display primary SMTP address
        (Get-Mailbox $uun | Format-List $output -ErrorAction Stop | Out-String).Trim()
        Write-Host "`n"
    }
    catch {
        # If for some reason there is an error then display message
        Write-Host "Unable to display!" -ForegroundColor Red
        Write-Host "Cause: $_." -ForegroundColor Red
        
        # Logout of Office 365
        Logout-OfOffice365
    }
}

# Function for checking current owner permissions - Requires an owner and a delegate user
function Check-CalendarPermissions {
    param(
        [Parameter(Mandatory=$true)][string]$calendarOwner,
        [Parameter(Mandatory=$true)][string]$delegateUser
    )
       
    # Get a list of users with permissions on the owners calendar
    $listOfUsers = Get-MailboxFolderPermission -Identity $calendarOwner@ed.ac.uk:\calendar
    # For each user in the list
    foreach ($entry in $listOfUsers){
        # Get the display name of the user requesting permission
        $userDisplayName = Get-User $delegateUser | Select-Object -ExpandProperty DisplayName
        # If the display name of the user requesting is equal to a user who currently has permissions
        if ($entry.User.DisplayName -eq "$userDisplayName") {
            # User already has some form of permissions on the owners calendar
            $permissionExists = "User already has permission"
            # Exit loop
            break           
        }
        # Else the user has no current permissions on the owners calendar
        else {
            $permissionExists = "User has no permission"                   
        }
    }
    # Return whether user has permissions or not
    return $permissionExists
}

# Function for displaying permissions on the owners calendar - needs the owners username
function Display-CalendarPermissions {
    param ([Parameter(Mandatory=$true)][string]$calendarOwner)

    Login-ToOffice365

    #Display Output header
    Write-Host "`n"    
    Write-Host "Current Calendar Permissions on $calendarOwner"
    Write-Host "______________________________________________"
    Write-Host "`n"
    # Display the users and what access rights they have on the owners calendar. Trim the blank lines at the end of the output
    (Get-MailboxFolderPermission -Identity $calendarOwner@ed.ac.uk:\calendar | select User,AccessRights | Format-Table -Auto | Out-String).Trim()
}

# Function for adding permission to owners calendar - needs the owners username and the username of the person requesting permisson 
function Add-CalendarPermission {
    param(
        [Parameter(Mandatory=$true)][string]$calendarOwner,
        [Parameter(Mandatory=$true)][string]$delegateUser
    )

    Login-ToOffice365

    # Begin loop
    do {
        [int[]] $validArray = 0..10
        # Ask what permission is to be applied for the user on the owners calendar
        Write-Host "`n"
        Write-Host "What permissions do you wish to give to $delegateUser on $calendarOwner's calendar ?"
        Write-Host "`n"
        Write-Host " 1 - Reviewer"
        Write-Host " 2 - Owner"
        Write-Host " 3 - Publishing Editor"
        Write-Host " 4 - Editor"
        Write-Host " 5 - Publishing Author"
        Write-Host " 6 - Author"
        Write-Host " 7 - Non Editing Author"
        Write-Host " 8 - Contributor"
        Write-Host " 9 - Limited Details (Free/Busy time, Subject, Location)"
        Write-Host "10 - Availability Only (Free/Busy time)"
        Write-Host "`n"
        Write-Host " 0 - Quitr"
        Write-Host "`n"
        $addPermissions = Read-Host "Enter Option"
        # Switch statement to deal with script tunner input
        switch($addPermissions) {
            # If input is 0, break from switch statement
            0 {$permission = 0;return}
            # If input is 1, set permission to Reviewer
            1 {$permission = "Reviewer"}
            # If input is 2, set permission to Owner
            2 {$permission = "Owner"}
            # If input is 3, set permission to PublishingEditor
            3 {$permission = "PublishingEditor"}
            # If input is 4, set permission to Editor
            4 {$permission = "Editor"}
            # If input is 5, set permission to PublishingAuthor
            5 {$permission = "PublishingAuthor"}
            # If input is 6, set permission to Author
            6 {$permission = "Author"}
            # If input is 7, set permission to NonEditingAuthor
            7 {$permission = "NonEditingAuthor"}
            # If input is 8, set permission to Contributor
            8 {$permission = "Contributor"}
            # If input is 9, set permission to LimitedDetails
            9 {$permission = "LimitedDetails"}
            # If input is 10, set permission to AvailabilityOnly
            10{$permission = "AvailabilityOnly"}   
        }
        # Check array to make sure user entry is between 1-10
        $validEntry = $validArray -Contains $addPermissions
        # If the user entry is not between 1-10 then display error
        if ($validEntry -eq $false) {
            Write-Host "`n*** ERROR *** - Invalid option. Please enter a number between 0 - 10.`n" -ForegroundColor Red
        }
        # else user entry is valid
        else {
            # Check if user already has permissions
            $delegatePermissionsExist = (Check-CalendarPermissions $calendarOwner $delegateUser)
            if ($delegatePermissionsExist -eq "User already has permission") {
                Write-Host "$delegateUser already seems to have the following permissions on $calendarOwner's calendar :" -ForegroundColor Yellow
                (Display-CalendarPermissions $calendarOwner)
                Write-Host "`nYou will need to remove these permissions and then re-add if you wish to change the current permissions for $delegateUser." -ForegroundColor Yellow                
                do {    
                    $question = "Do you wish to remove current permissions for $delegateUser on $calendarOwner's calendar (y/n) ?"
                    $confirmReply = (Confirm-Answer $question)
                } until ($confirmReply -ne "error")
            }
            

            try {
                # Add the permission and supress output
                Add-MailboxFolderPermission -Identity $calendarOwner@ed.ac.uk:\calendar -User $delegateUser -AccessRights $permission -ErrorAction Stop -Confirm:$false > $null
                # Display message
                Write-Host "`n"
                Write-Host "*** $permission rights have been successfully added for $delegateUser on $calendarOwner's calendar ***" -ForegroundColor Green
                Write-Host "`n"
            }
            # If adding permission fails
            catch {
                # Display error message
                Write-Host "`n"
                Write-Host "*** ERROR *** - Unable to add rights!" -ForegroundColor Red
                Write-Host "Cause : $_." -ForegroundColor Red
            }   
        }               
    # Loop until input is greater than or equal to 0 or until input is less than or equal to 10                      
    } until ($validEntry -eq $true)        
}

# Function for removing permissions - needs the owners username and the username of the person who's permission will be removed 
function Remove-CalendarPermission {
    param (
        [Parameter(Mandatory=$true)][string]$calendarOwner,
        [Parameter(Mandatory=$true)][string]$delegateUser
    )
    
    Login-ToOffice365

    # Begin loop
    do {
        # By the time this function is called, all checks have been done
        # Display what rights the user currently has on the owners calendar
        try {
            Get-MailboxFolderPermission -Identity $calendarOwner@ed.ac.uk:\Calendar -User $delegateUser -ErrorAction Stop | select User,AccessRights | Format-Table -Auto | Out-Host
        }
        # Display error message if the user does not currently have rights on the calendar
        catch {
            Write-Host "`n"
            Write-Host "*** ERROR *** A problem has occurred." -ForegroundColor Red
            Write-Host "Cause : $_."
            Write-Host "`n"
            break
        }
        Write-Host "$delegateUser currently has permissions listed above on $calendarOwner's calendar.`n"
        $permission = (Get-MailboxFolderPermission -Identity $calendarOwner@ed.ac.uk:\Calendar -User $delegateUser).AccessRights
        # Ask for confirmation
        $removePermissions = Read-Host "Do you wish to remove all permissions on $calendarOwner's calendar for $delegateUser (y or n) ? "
        # Convert input to upper case
        $removePermissions = $removePermissions.ToUpper()
        # Begin switch statement
        switch ($removePermissions){
            # If input equals Y
            Y {
                # Attempt to remove permision
                try {
                    # Remove permission and suppress output
                    Remove-MailboxFolderPermission -Identity $calendarOwner@ed.ac.uk:\calendar -User $delegateUser -ErrorAction Stop -Confirm:$false
                    # Display message
                    Write-Host "`n"
                    Write-Host "*** $permission rights for $delegateUser have been successfully removed from $calendarOwner's calendar ***" -ForegroundColor Green
                    Write-Host "`n"
                }
                # If there is a problem removing the permission
                catch {
                    # Display error message
                    Write-Host "`nUnable to remove $permission rights!" -ForegroundColor Red
                    Write-Host "Cause : $_."
                }
            }
            # If input equals N
            # Display message and break from switch statement
            N {Write-Host "`nPermissions have not been removed." -ForegroundColor Yellow;break}
            # If an invalid option is chosen, display error message
            default {
                Write-Host "`n"
                Write-Host "*** Invalid option. Please enter Y or N ***" -ForegroundColor red
                Write-Host "`n"    
            }
        }
    # Loop until input equals Y or N   
    } until ($removePermissions -eq "Y" -or $removePermissions -eq "N")
}


function Set-CalendarPermissions {
    
    # Make sure we are logged
    Login-ToOffice365

    # Loop until valid uun is supplied
    do {        
        $calendarOwner = Read-Host "Please enter username of calendar owner (or CTRL+C to quit)"
        # Convert username to upper case
        $calendarOwner = $calendarOwner.ToLower()
        $calendarOwnerExists = (Check-Mailbox $calendarOwner) 
    } until($calendarOwnerExists -eq "User exists")

    Write-Host "`n"
    # Begin nested loop for owner options and get input
    do {   
        Write-Host "What would you like to do on account $calendarOwner ?"
        Write-Host "`n"
        Write-Host "1 - List current permissions"
        Write-Host "2 - Add access rights for another user"
        Write-Host "3 - Remove access rights for another user"
        Write-Host "`n"
        Write-Host "0 - Quit"
        Write-Host "`n"
        $menuChoice = Read-Host -Prompt "Please enter an option for account $owner"    
        switch ($menuChoice) {
            1 {(Display-CalendarPermissions $calendarOwner)} # Display owners permissions
            2 { do {        
                    $delegateUser = Read-Host "Please enter username of calendar owner (or CTRL+C to quit)"
                    # Convert username to lower case
                    $delegateUser = $delegateUser.ToLower()
                    $delegateUserExists = (Check-Mailbox $delegateUser) 
                } until($delegateUserExists -eq "User exists")
                
                # If the account exists, check to see if the user currently has any rights on the owners calendar
                $currentPermissions = (Check-CalendarPermissions $calendarOwner $delegateUser)
                if ($currentPermissions -eq "User already has permission"){
                    Write-Host "`n"
                    Write-Host "$delegateUser already has permission of some kind on $calendarOwner's calendar." -ForegroundColor Yellow
                    Write-Host "If you wish to change these permissions please remove the current permissions and then re-add them with the new permissions." -Foreground Yellow
                    Write-Host "`n"
                    Quit-Process                       
                }

                # Once we have a valid username and all checks are done, grant the user permission to the owners account
                (Add-CalendarPermission $calendarOwner $delegateUser)                    
            }
            3 { do {        
                    $delegateUser = Read-Host "Please enter username of delegate (or CTRL+C to quit)"
                    # Convert username to lower case
                    $delegateUser = $delegateUser.ToLower()
                    $delegateUserExists = (Check-Mailbox $delegateUser) 
                } until($delegateUserExists -eq "User exists")
                        
                $currentPermissions = (Check-CalendarPermissions $calendarOwner $delegateUser)
                if ($currentPermissions -eq "User has no permission"){
                    Write-Host "$delegateUser currently has no permissions on $calendarOwner's calendar" -ForegroundColor Yellow
                    Write-Host "`n"
                    Quit-Process
                }
                  
                # Once all checks have been passed, remove the permissions
                (Remove-CalendarPermission $calendarOwner $delegateUser)    
            }
            # Break out of loop
            0 {Quit-Process}
            # Display error message if user has entered an invalid option
            default {                
                Write-Host "`n*** Invalid option. Please enter either 0, 1, 2 or 3 ***`n" -ForegroundColor red                     
            }            
        }                        
    } until ($menuChoice -eq 0)
}