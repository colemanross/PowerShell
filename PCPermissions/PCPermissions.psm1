
# Known issues with *-LocalGroupMember commands. For just now, using older trusted commands for previous PowerShell Versions
# https://github.com/PowerShell/PowerShell/issues/2996
# https://superuser.com/questions/1131901/get-localgroupmember-generates-error-for-administrators-group
# Function for displaying local permissions. This will only work on PS versions 5 and later.
function Display-LocalGroupMembers{
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    Invoke-Command -ComputerName $PCName -ScriptBlock {
        param($PCName, $groupName)
        try{
            $members = (Get-LocalGroupMember -Group $groupName).Name
            Write-Host "List of $groupName on $PCName"
            Write-Host "========================================"
            Foreach ($member in $members) {
                if ($member -eq "ED\" -or $member -eq "$PCName\Administrator"){
                    continue
                }
                Write-Host $member
            }
        }
        catch {
            Write-Host "`nUnable to obtain members of $groupName on $PCName! Cause : $_.`n" -ForegroundColor Red
        }
    } -ArgumentList $PCName, $groupName
}

# Function for adding a member to a local group. This will only work on PS versions 5 and later
function Add-MemberToLocalGroup{
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    
    Invoke-Command -ComputerName $PCName -ScriptBlock {
        param($PCName, $groupName)        
        $group = "ED\$PCName"
        try{
            $members = (Get-LocalGroupMember -Group $groupName).Name
            if ($group -in $members){            
                $question = "$group already exists in $groupName. Do you wish to remove and re-add (y/n) ?"
                $answer = ${function:Confirm-Answer $question}
                if ($answer -eq "n"){
                    return
                }
                else{
                    try{
                        (Delete-LocalGroupMember $PCName $groupName)
                        Write-Host "Successfully removed $PCName from $groupName. Attempting to re-add..."
                    }
                    catch{
                        Write-Host "`nUnable to remove $PCName from $groupName! Cause : $_.`n" -ForegroundColor Red
                    }
                }
            }
            else{
                try{
                    Add-LocalGroupMember -Group "$groupName" -Member "$group"
                    Write-Host "`nSuccessfully added $group to local $groupName group on $PCName.`n" -ForegroundColor Green 
                }
                catch{
                    Write-Host "`nUnable to add $PCName to local $groupName group! Cause : $_.`n" -ForegroundColor Red
                }
                    
                
             }  
            
        }
        catch{
            Write-Host "`nUnable to add $group to local $groupName group! Cause : $_.`n" -ForegroundColor Red
        }
        
    } -ArgumentList $PCName, $groupName
}

# Function for removing a member from a local group. This will only work on PS versions 5 and later.
function Delete-MemberFromLocalGroup{
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    Invoke-Command -ComputerName $PCName -ScriptBlock {
        param($PCName, $groupName)
        $group = "ED\$PCName"
        try{
            $members = (Get-LocalGroupMember -Group $groupName).Name
            if ($group -notin $members){
                Write-Host "`n$group does not appear to be in the local $groupName group!`n" -ForegroundColor Red
            }
            else{
                try{
                    Remove-LocalGroupMember "Administrators" -Member $group
                }
                catch{
                    Write-Host "`nUnable to remove $group from local $groupName group! Cause : $_.`n" -ForegroundColor Red
                }
            }
        }
        catch{
            Write-Host "`nUnable to remove $group from local $groupName group! Cause : $_.`n" -ForegroundColor Red
        }
        
    } -ArgumentList $PCName, $groupName
}

# Function for displaying members of a local group. This should work on any PS versions
function Display-LocalGroupMembersWin7{
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    # Get group from specified computer
    $group = ([ADSI]"WinNT://$PCName/$groupName")
    # Get current members
    $tempMembers = @($group.psbase.Invoke("Members"))
    # Store each member of the group 
    $members = $tempMembers | foreach {$_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null)} # Store each member of the group    
    # Display Users within group
    Write-Host "`nList of $groupName on $PCName"
    Write-Host "============================================"
    foreach ($member in $members){
        Write-Host $member
    }
}

# Function to add AD group to local PC group
function Add-ADGroupToLocalGroupWin7 {
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    # Get the local group. Display error if there is a problem.
    try {    
        $localGroup = [ADSI]"WinNT://$PCName/$groupName,group" # Get group from specified computer
    }
    catch {
        Write-Host "`nUnable to obtain information from $PCName. Cause $_."
        return
    }
    # Get current members
    $tempMembers = @($localGroup.psbase.Invoke("Members")) 
    # Store each member of the group
    $members = $tempMembers | foreach {$_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null)}
    # If AD Group is already in the local group
    if ($PCName -in $members) {
        Write-Host "`nAD Group $PCName already appears to be in the local $groupName group on PC $PCName!`n" -ForegroundColor Yellow
        return
    }
    # Else, attempt to add AD group
    else {
        Write-Host "`nAttempting to add AD group $PCName to local $groupName on $PCName..."
        # Try adding the group
        try {
            $localGroup.Add("WinNT://ed.ac.uk/$PCName,group") #Add the AD Group to the local group      
            Write-Host "`n$PCName has been successfully added to the $groupName group on PC $PCName`n" -ForegroundColor Green # Display successfull message 
        }
        catch {
            Write-Host "`nUnable to add $PCName to $groupName group on PC $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }
    }         
}

# Function to add AD user to local PC group
function Add-ADUserToLocalGroupWin7 {
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )

    # Get username to add. Loop until a valid uun is supplied
    do {
        $ADUser = Read-Host "Enter AD username to add to local $groupName group "              
        $userExists = (Check-ADuser $ADUser)
    } until ($userExists -eq "User exists")
    # Try adding the AD user to the local group
    try {    
        $localAdminGroup = [ADSI]"WinNT://$PCName/$groupName,group" # Get group from specified computer        
        $localAdminGroup.Add("WinNT://ed.ac.uk/$ADUser,user") #Add the username to the group  
        Write-Host "`n$ADUser has been successfully added to the local $groupName group on $PCName`n" -ForegroundColor Green # Display successfull message         
    } 
    catch {
        Write-Host "`nUnable to add $ADUser to the local $groupName on PC $PCName!" -ForegroundColor Red
        Write-Host "Cause :" $_. -ForegroundColor Red
    }  
}

# Function to remove AD group from local PC group
function Remove-ADGroupFromLocalGroupWin7 {
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    # Get group from specified computer. Display error if needed.
    try {    
        $localGroup = [ADSI]"WinNT://$PCName/$groupName,group" 
    }
    catch {
        Write-Host "`nUnable to obtain information from $PCName." -ForegroundColor Red
        Write-Host "Cause $_.`n" -ForegroundColor Red
        return
    }
    # Get current members
    $tempMembers = @($localGroup.psbase.Invoke("Members"))
    # Store each member of the group 
    $members = $tempMembers | foreach {$_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null)}
    # Check if group is currently present. If not display error. 
    if ($PCName -notin $members) {
        Write-Host "$PCName doesn't appear to currently be in the local $groupName group on PC $PCName!" -ForegroundColor Yellow
    }
    # Else, if the group is present
    else {
        Write-Host "`nAttempting to remove AD Group $PCName from local $groupName on PC $PCName...."
        # Attempt to remove
        try {
            ([ADSI]"WinNT://$PCName/$groupName,group").remove("WinNT://ed.ac.uk/$PCName") # Remove group
            Write-Host "`n$PCName has been successfully removed from the $groupName group on $PCName!`n" -ForegroundColor Green 
        }
        catch {
            Write-Host "`nUnable to remove $PCName from local $groupName group on PC $PCName!" -ForegroundColor Red
            Write-Host "Cause :" $_. -ForegroundColor Red
        }
    }    
}

# Function to remove AD user from local group
function Remove-ADUserFromLocalGroupWin7 {
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )

    $ADUser = Read-Host -Prompt "Please enter the username to be removed from the $groupName Group" # Get username
    do {              
        $userExists = (Check-ADuser $ADUser)
    } until ($userExists -eq $true)

    try {    
        $localMembers = [ADSI]"WinNT://$PCName/$groupName,group" # Get group from specified computer
    }
    catch {
        Write-Host "`nUnable to obtain information from $PCName. Cause $_."
        return
    }
    $tempMembers = @($localMembers.psbase.Invoke("Members")) # Get current members
    $members = $tempMembers | foreach {$_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null)} # Store each member of the group
    if ($ADUser -notin $members){
        Write-Host "`n$ADUser doesn't appear to currently be in the local $groupName group on PC $PCName!`n" -ForegroundColor Yellow
    }
    else {
        Write-Host "`nAttempting to remove $adUser from local $groupName on PC $PCName...."
        try {
            ([ADSI]"WinNT://$PCName/$groupName,group").remove("WinNT://ed.ac.uk/$ADUser") # Remove user from group
            Write-Host "`n$ADUser has been successfully removed from the $groupName group on $PCName!`n" -ForegroundColor Green
        }
        catch {
            Write-Host "`nUnable to remove $ADUser from local $groupName group on PC $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        } 
    }
}