function Display-ADGroupMembers{
    <#
        .Synopsis
        Displays members of an AD group.

        .Description
        Displays members of an AD group.

        REQUIREMENTS:
        - AD read access

        WHAT IT DOES:
        - Attempts to run Get-ADGroupMember

    #>

    param([Parameter(Mandatory=$true)][string]$groupName)
    # Get existing users in the group
    try{
        $existingUsers = (Get-ADGroupMember $groupName -ErrorAction SilentlyContinue).SamAccountName
    }
    catch {
        Write-Host "`nUnable to obtain members of $groupName!" -ForegroundColor Red
        Write-Host "Cause :$_.`n" -ForegroundColor Red
    }
    # For each user in the group
    Write-Host "`nCurrent users in $groupName"
    Write-Host "------------------------------"
    Write-Host "UUN`t`tFULL NAME"
    Write-Host "---`t`t--------"
    foreach ($user in $existingUsers) {
        $fullName = (Return-FullName $user)
        Write-Host "$user`t$fullName"
    }    
}

function Display-FullName {
    <#
        .Synopsis
        Displays full name of user.

        .Description
        Displays full name of the user in "SURNAME Forename" format.

        REQUIREMENTS:
        - AD read access

        WHAT IT DOES:
        - Checks to make sure uun exists in AD
        - Runs "Get-ADUser" and displays user in SURNAME Forename format

    #>
   param ([Parameter(Mandatory=$true)][string]$uun)

   $userExists = (Check-ADuser $uun)
   if ($userExists -eq "error") {
        return
    }
    else {
        Get-ADUser $uun -Properties DisplayName | Select-Object DisplayName   
    }

}

function Return-FullName {
    <#
        .Synopsis
        Returns full name of user instead of displaying.

        .Description
        Returns full name of user instead of displaying. Is returned in "SURNAME Forename" format. 
        
        REQUIREMENTS:
        - AD read access.
        
        WHAT IT DOES:
        - Checks to make sure uun exists in AD
        - Returns full name of uun in format SURNAME Forename instad of displaying on screen  

    #>
    param ([Parameter(Mandatory=$true)][string]$uun)

    $userExists = (Check-ADuser $uun)
    if ($userExists -eq "error") {
        return
    }
    else {
        $fullName = Get-ADUser $uun -Properties DisplayName | Select-Object DisplayName
        return $fullName.DisplayName
    }
}

function return-FirstName {
    param ([Parameter(Mandatory=$true)][string]$uun)

    # Get full name from AD
    $fullName = (Return-FullName $uun)
    # Separate 1st, 2nd, 3rd etc.. elements of name into an array
    $firstName_ = $fullName.split()
    # Take last array element as first name
    $firstName = $firstName_[-1]
    # Return first nameHang
    return $firstName
}

function Display-DisabledPCs {
    <#
        .Synopsis
        Displays disabled PC objects in AD.

        .Description
        Displays disabled PC objects in AD within the following OUs :
            - ou=LF,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk
            - ou=CCBS,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk
            - ou=SBMS,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk
            - ou=Central,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk
            - ou=CPHS,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk
            - ou=MVMCOLLEGE,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk
            - ou=BQ,ou=MVM,ou=ISSD,ou=SDX,ou=UoEX,dc=ed,dc=ac,dc=uk

        REQUIREMENTS:
        - Read permissions on AD     
        
        WHAT IT DOES:
        - Accepts array of PC netBIOS Names
        - Sorts the array by name
        - Iterates through array of PC names and outputs disabled PC objects  

    #>

    # Nested function to display machines
    function Display-Computers{
        param ([Parameter(Mandatory=$true)][array]$PCs)

        $sortedPCs = $PCs.Name | sort
        foreach($PC in $sortedPCs){
            Write-Host $PC
        }
    }

    # Non Windows 10
    Write-Host "NON-WINDOWS 10 MACHINES :"
    Write-Host "LF"
    Write-Host "==========="
    $PCsLF = Get-ADComputer -Filter * -SearchBase "ou=LF,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCsLF)
    $LF = $PCsLF.Count
    Write-Host "`nCCBS"
    Write-Host "==========="
    $PCsCCBS = Get-ADComputer -Filter * -SearchBase "ou=CCBS,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCsCCBS)
    $CCBS = $PCsCCBS.Count
    Write-Host "`nSBMS"
    Write-Host "==========="
    $PCsSBMS = Get-ADComputer -Filter * -SearchBase "ou=SBMS,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCsSBMS)
    $SBMS = $PcsSBMS.Count
    Write-Host "`nCentral"
    Write-Host "==========="
    $PCsCentral = Get-ADComputer -Filter * -SearchBase "ou=Central,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCsCentral)
    $Central = $PCsCentral.Count
    Write-Host "`nCPHS"
    Write-Host "==========="
    $PCsCPHS = Get-ADComputer -Filter * -SearchBase "ou=CPHS,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCsCPHS)
    $CPHS = $PCsCPHS.Count
    Write-Host "`nMVMCOLLEGE"
    Write-Host "================="
    $PCsCollege = Get-ADComputer -Filter * -SearchBase "ou=MVMCOLLEGE,ou=MVM,ou=ISD,ou=SD7,ou=UoESD,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCsCollege)
    $College = $PCsCollege.Count

    $total = $LF + $CCBS + $SBMS + $Central + $CPHS + $College
    Write-Host "`nTotal : $total`n"

    # Windows 10
    Write-Host "WINDOWS 10 MACHINES :"
    Write-Host "==============================="
    $PCs = Get-ADComputer -Filter * -SearchBase "ou=BQ,ou=MVM,ou=ISSD,ou=SDX,ou=UoEX,dc=ed,dc=ac,dc=uk" | Where-Object{$_.Enabled -eq $False} | Select-Object Name
    (Display-Computers $PCs)
    $PCsWin10 = $PCs.Count
    Write-Host "`nTotal Win 10 machine disabled : $PCsWin10`n"
}

function Display-ADInfo {
    <#
        .Synopsis
        Displays AD info on PC object.

        .Description
        Displays the following info on PC AD object :
        - DistinguishedName (OU Path)
        - DNS Hostname
        - Enabledd / disabled
        - Object class
        - Object GUID
        - SAMAccountName (Security Accounts Manager)
        - SID (Security Identifier)
        - UserPrincipalName (mainly used for usernames, not PCs)

        REQUIREMENTS:
        - Read permissions on AD

        WHAT IT DOES:
        - Runs "Get-ADComputer" on specific PC
       
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    Get-ADComputer -Identity $PCName
}

# Function to display OU Path of PC object
function Display-PCOUPath {
    <#
        .Synopsis
        Displays OU path for a PC.

        .Description
        Displays OU path for a PC.

        REQUIREMENTS:
        - Read permissions on AD

        WHAT IT DOES:
        - Runs "Get-ADComputer" on specific PC but only grabs Distinguished name (OU Path)
        - Cut's the machine name from the end of the path
        - Displays the path       
    #>
    param ([Parameter(Mandatory=$true)][string]$PCName)

    $OU_ = (Get-ADComputer -Identity $PCName).DistinguishedName
    $OU = ($OU_ -split ",",2)[1]

    Write-Host "`nOU path for $PCName :"
    Write-Host "$OU`n"
}

# Function to return OU path of PC object
function Return-PCOUPath {
    <#
        .Synopsis
        Returns an OU path for a PC.

        .Description
        Returns an OU path for a PC.

        REQUIREMENTS:
        - Read permissions on AD

        WHAT IT DOES:
        - Runs "Get-ADComputer" on specific PC but only grabs Distinguished name (OU Path)
        - Cut's the machine name from the end of the path
        - returns the OU path to the caller       
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    $OU_ = (Get-ADComputer -Identity $PCName).DistinguishedName
    $OU = ($OU_ -split ",",2)[1]
    return $OU
}

# Function to copy OU path of PC object to clipboard
function Copy-PCOUPath {
    <#
        .Synopsis
        Copies OU path for a PC.

        .Description
        Copies OU path for a PC.

        REQUIREMENTS:
        - Read permissions on AD

        WHAT IT DOES:
        - Runs "Get-ADComputer" on specific PC but only grabs Distinguished name (OU Path)
        - Cut's the machine name from the end of the path
        - Copies the OU path to the clipboard       
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    $OU_ = (Get-ADComputer -Identity $PCName).DistinguishedName
    $OU = ($OU_ -split ",",2)[1]

    Set-Clipboard -Value $OU
    Write-Host "`n$OU copied to clipboard!`n"

}

function Return-GroupOUPath {
    param ([Parameter(Mandatory=$true)][string]$GroupName)

    $OU_ = (Get-ADGroup -Identity $GroupName).DistinguishedName
    $OU = ($OU_ -split ",",2)[1]
    return $OU
}

# Function to check if machine resides in an MDSD group
function Check-MDSD($PCName, $searchString) {
    # Get list of groups machine resides in.
    $groupList = Get-ADComputer $PCName | Get-ADPrincipalGroupMembership | Select Name
    # For each item in the list, check to see if MDSD is in the name
    foreach ($item in $groupList){
        # Convert custom object to string
        $itemString = ($item | Out-String).Trim()
        if ($itemString -like $searchString){
            # Return message and break from function
            return "$PCName also appears to be a laptop as it is a member of the following AD Group : `n$itemString "
            break
        }
        # Else continue with the next group
        else {
            continue
        }
    }
}

function Display-PCMemberOf {
    param([Parameter(Mandatory=$true)][string]$PCName)

    $allGroups = New-Object System.Collections.ArrayList
    
    $memberOf_ = (Get-ADComputer $PCName -Properties *).MemberOf
    foreach($group in $memberOf_){
        $CN_ = ($group -split ",",2)[0]
        $CN = $CN_.TrimStart("CN=")
        $allGroups.Add($CN) > $null
    }
    Write-Host "`n$PCName is a member of :"
    $output = $allGroups -join "`-n"
    foreach ($group in $allGroups){
        Write-Host "`t$group"
    }
}

function Return-PCMemberOf {
    param([Parameter(Mandatory=$true)][string]$PCName)

    $allGroups = New-Object System.Collections.ArrayList
    
    $memberOf_ = (Get-ADComputer $PCName -Properties *).MemberOf
    foreach($group in $memberOf_){
        $CN_ = ($group -split ",",2)[0]
        $CN = $CN_.TrimStart("CN=")
        $allGroups.Add($CN) > $null
    }

    return [System.Collections.ArrayList]$allGroups
}

function Display-UserMemberOf ([string]$uun){
    
}

