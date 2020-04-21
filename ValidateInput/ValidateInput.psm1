function Confirm-Answer{
    param([Parameter(Mandatory=$true)][string]$question)

    $answer = Read-Host $question
    $answer = $answer.ToLower()
    if ($answer -eq "y"){
        return "y"
    }
    elseif ($answer -eq "n"){
        return "n"
    }

    else {
        Write-Host "`n*** Invalid entry! *** - Please enter y or n !`n" -ForegroundColor Red
        return "error"
    }
}

function Check-ADUser {
    param ([Parameter(Mandatory=$true)][string]$userName)

    try{
        If ([bool] (Get-ADUser "$userName" -ErrorAction SilentlyContinue)) {       
            return "User exists"
        }
    }      
    
    catch {
        Write-Host "`n$userName cannot be found in AD!`n" -ForegroundColor Red
        return "error"
    }
}

function Check-ADGroup {
    param ([Parameter(Mandatory=$true)][string]$groupName)

    try {
        if ([bool] (Get-ADGroup "$groupName" -ErrorAction SilentlyContinue)) {
            return "Group exists"
        }
    }
    catch {
        Write-Host "`nAD group $groupName cannot be found!`n" -ForegroundColor Yellow
        return "error"
    }
}

function Check-ADComputer {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    try{
        if ([bool] (Get-ADComputer "$PCName" -ErrorAction SilentlyContinue)) {
            return "PC exists"
        }
    
    }
    catch {
        Write-Host "`n$PCName cannot be found in AD!`n" -ForegroundColor Yellow
        return "error"
    }
}

function Check-PCisOnline {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    try {
        if (Test-Connection $PCName -count 1 -ErrorAction Stop) {
            return "PC online"
        }
    } 
    catch {
        Write-Host "`n$PCName is offline!`n" -ForegroundColor Red
        return "error"
    }
}

function Check-UserInGroup {
    param(
        [Parameter(Mandatory=$true)][string]$groupName,
        [Parameter(Mandatory=$true)][string]$userName
    )

    # Get existing users in the group
    try{
        $existingUsers = (Get-ADGroupMember $groupName -ErrorAction SilentlyContinue).SamAccountName
    }
    catch {
        Write-Host "`nUnable to obtain members of $groupName!" -ForegroundColor Red
        Write-Host "Cause :" $_. -ForegroundColor Red
    }
    
    if ($userName -in $existingUsers){
        return "User in group"
    }
    else {
        return "User not in group"
    }
}

function Check-Mailbox {
    param ([Parameter(Mandatory=$true)][string]$uun)

    try{
        If ([bool] (Get-Mailbox "$uun" -ErrorAction SilentlyContinue)) {       
            return "User exists"
        }
    }      
    
    catch {
        Write-Host "`n$uun cannot be found in Office365!`n" -ForegroundColor Red
        return "error"
    }
}