function Add-ADUserToADGroup {
    <#
        .Synopsis
        Adds an AD uun to an AD group.

        .Description
        Adds an AD uun to an AD group.

        REQUIREMENTS:
        - AD read  & write access

        WHAT IT DOES:
        - Obtains sAMAccountName from the supplied parameter
        - Asks for uun
        - Checks to make sure uun is valid
        - Checks to see if uun already exists in AD group
        - If not already in group, attempts to add uun to group

    #>

    param(
        [Parameter(Mandatory=$true)][object]$groupName,
        [Parameter(Mandatory=$true)][string]$groupType
    )

    $cnName = (Get-ADGroup $groupName).SamAccountName
  
    do {
        # Get username to be added to group
        $userName = Read-Host "`nPlease enter username to add to AD $groupType group $cnName (or CTRL+C to quit)"
        # Convert username to lowercase
        $userName = $username.ToLower()                   
        $userExists = (Check-ADUser $userName)            
    # Loop until username exists
    } until ($userExists -eq "User exists")


    $doesExist = (Check-UserInGroup $cnName $userName)
    # If not in group
    if ($doesExist -eq "User not in group") {
        # Attempt to add username to the group and display message
        try {
            (Add-ADGroupMember $cnName $userName)
            Write-Host "`nAD user $userName added to AD group $cnName`n" -ForegroundColor Green
        }
        # If for any reason there is an error when attempting to add, display message
        catch {
            Write-Host "`nUnable to add AD user $userName to AD Group $cnName." -ForegroundColor Red
            Write-Host "Cause : $_.`n" -ForegroundColor Red
        }
    }
    # Else username is already in group
    else {Write-Host "`nAD user $userName already appears to be in AD group $cnName`n" -ForegroundColor Yellow}

    return $username
}

function Remove-ADUserFromADGroup {
     <#
        .Synopsis
        Removes a uun from an AD group.

        .Description
        Removes a uun from an AD group.

        REQUIREMENTS:
        - AD read  & write access

        WHAT IT DOES:
        - Asks for uun
        - Checks to make sure uun is valid
        - Checks to see if uun already exists in AD group
        - If in group, attempts to remove uun from group

    #>

    param([Parameter(Mandatory=$true)][string]$groupName)

    do {
        # Get username to be added to group
        $userName = Read-Host "`nPlease enter username to remove from group $groupName (or CTRL+C to quit)"
        # Convert username to lowercase
        $userName = $username.ToLower()
        # If username equals q
        if ($username -eq "q") {
            # return
            return
        }
        else {           
            $userExists = (Check-ADUser $userName)            
        }
        # Loop until a valid user is entered
    } until ($userExists -eq "User exists")

    # Check to see if username is currently in group
    $doesExist = (Check-UserInGroup $groupName $userName)
    # If not in group
    if ($doesExist -eq "User in group") {
        # Attempt to remove and display message
        try {
            Remove-ADGroupMember $groupName $userName -Confirm:$false
            Write-Host "`nAD user $userName successfully removed from AD group $groupName!`n" -ForegroundColor Green
        }
        # If there is an error when trying to remove then display message
        catch {
            Write-Host "Unable to remove AD user $userName from AD group $groupName." -ForegroundColor red
            Write-Host "Cause :$_.`n" -ForegroundColor Red        
        }  
    }
    # Else the user is not in the group
    else {Write-Host "`nAD user $userName is not in AD group $groupName!`n" -ForegroundColor Red}    
}

function Remove-PCFromADGroup{
    <#
        .Synopsis
        Removes a PC Object from an AD group.

        .Description
        Removes a uun from an AD group.

        REQUIREMENTS:
        Removes a PC Object from an AD group.

        WHAT IT DOES:
        - Asks for uun
        - Checks to make sure uun is valid
        - Checks to see if uun already exists in AD group
        - If in group, attempts to remove uun from group

    #>
    # Loop until valid PC name is entered (or script is quit)
    do {        
        $PCName = Read-Host "Please enter name of PC to remove from AD Group (or q to return)"
        # Convert PC to uppercase
        $PCName = $PCName.ToUpper()
        #If PC equals q
        if ($PCName -eq "Q") {
            Write-Host ""
            return
        }
        else {
            $PCExists = (Check-ADComputer $PCName)
        }
    } until ($PCExists -eq "PC exists")    
    
    # Get list of AD groups $PCName belongs to
    $memberOf = Return-PCMemberOf $PCName
    
    if ($memberOf.count -eq 0){
        Write-Host "`nPC AD object $PCName does not appear to be a member of any AD groups`n" -ForegroundColor Yellow
        return
    }

    # Set index reference to 1
    $index = 1
    # Display current member of
    Write-Host "`n$PCName is currently a member of the following groups :"
    foreach ($group in $memberOf){        
        Write-Host "`t$index. $group"
        $index++
    }
    # If there is only 1 group, ask if PC is to be removed
    $count = $memberOf.count
    if ($count -eq 1){
        do {
            $question = "`n$PCName only appears to be a member of one AD Group - $memberOf.`nDo you wish to remove $PCName from $memberOf (y/n)?"
            $confirm = (Confirm-Answer $question)
        } until ($confirm -ne "error")
        $groupToRemove = $memberOf
    }
    # Else ask for group
    else {
        do {
            $remove = Read-Host "`nPlease enter number to remove (or 0 to quit) "
            # If reply is out of range display error message
            if($remove -eq 0){
                return
            }
            if ($remove -lt 1 -or $remove -gt $count) {            
                Write-Host "`nInvalid entry! Please enter a number from 1 - $count!" -ForegroundColor Red
            } 
        } until ($remove -ge 1 -and $remove -le $count)
    
        # Adjust for correct index reference
        $entryToremove = $remove -1
        # Get the group name
        $groupToRemove = $memberOf[$entryToremove]
    }
    
    # Try and remove the PC from the group
    try{
        Remove-ADGroupMember -Identity $groupToRemove -Members $PCName$ -Confirm:$false >$null
        Write-Host "`nSuccessfully removed $PCName as a member of $groupToRemove`n" -ForegroundColor Green
    }
    catch {
        Write-Host "`nUnable to remove $PCName as a member of $groupToRemove" -ForegroundColor Red
        Write-Host "Cause: $_.`n"
    }        
}