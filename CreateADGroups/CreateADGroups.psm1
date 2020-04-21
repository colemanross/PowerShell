# Function to create Win 7 / Win 8 group
function Create-Win7Group {    
    param(
    [Parameter(Mandatory=$true)][string]$PCName,
    [Parameter(Mandatory=$true)][string]$groupType
)
    # Ask what OU to create group
    do {
        Write-Host "`nWhat OU group do you wish to create the $groupType group in? `n "
        Write-Host " 1 - LF"
        Write-Host " 2 - CPHS"
        Write-Host " 3 - CCBS"
        Write-Host " 4 - CRFR"
        Write-Host " 5 - CRIC"
        Write-Host " 6 - SBMS"
        Write-Host " 7 - MVMCOLLEGE"
        Write-Host " 8 - BRR"
        Write-Host " 9 - CCNS`n"
        $targetOU = Read-Host "Select what OU to create the group in for $PCName (or q to change pc name) "
        $targetOUUpperCase = $targetOU.ToUpper()
            if ($targetOUUpperCase -eq "Q") {exit}
            # Begin switch statement
            switch ($targetOU) {
                1 {$subOU = "LF"}
                2 {$subOU = "CPHS"}
                3 {$subOU = "CCBS"}
                4 {$subOU = "CRFR"}
                5 {$subOU = "CRIC"}
                6 {$subOU = "SBMS"}
                7 {$subOU = "MVMCOLLEGE"}
                8 {$subOU = "BRR"}
                9 {$subOU = "CCNS"}
                default {Write-host "`nPlease enter a number from 1 - 9 (or q to change PC name) " -ForegroundColor Red}
            }
        # Loop until a valid entry is entered 
    } until ($targetOU -ge 1 -and $targetOU -le 7 -or $targetOU -eq "Q")

    # If reply is equal to Q then break
    if ($targetOUUpperCase -eq "Q") {exit}

    $description = Read-Host "`nPlease enter a description for the group "

    # Set samAccountName depending on group type
    if ($groupType -eq "Remote Users"){
        $samAccountName = "$pcName-rdp"
    }
    elseif ($groupType -eq "Administrators"){
        $samAccountName = "$pcName-adm"
    }
    else {
        Write-Host "`nUnable to determine group type!`n" -ForegroundColor Red
        return
    }

    # Set OU path
    $OU = "OU=$groupType,OU=$subOU,OU=Authorisation,OU=UoESD,DC=ed,dc=ac,dc=uk"
    # Attempt to create the group
    try {
        Write-Host "`nAttempting to create $groupType group....`n"    
        New-ADGroup -Name $pcName -SamAccountName $samAccountName -GroupCategory Security -GroupScope Global -Path $OU -Description $description -ErrorAction SilentlyContinue  
        Write-host "$groupType group for $pcName created at the following location :`n" -ForegroundColor Green
        Write-Host "$OU`n" -ForegroundColor Green
    }
    catch {
        Write-Host "`nUnable to create group!" -ForegroundColor Red
        Write-Host "Cause :$_.`n" -ForegroundColor Red
    }
}

# Function to create Windows 10 group
function Create-Win10Group {
    param(
        [Parameter(Mandatory=$true)][string]$pcName,
        [Parameter(Mandatory=$true)][string]$groupType
    )

    # Set OU path
    $OU = "ou=$groupType,ou=MVM,ou=Auth,ou=UoEX,dc=ed,dc=ac,dc=uk"
    # Set Description
    $description = Read-Host "`nPlease enter a description for the group - normally building, user and [P] for permanent (admin only) "
    # Set SamAccountName
    $samAccountName = $pcName

    # If the group is a remote desktop group then we need to append "-rdp"
    if ($groupType -eq "RemoteUsers"){
        $OU = "ou=RemoteUsers,ou=MVM,ou=Auth,ou=UoEX,dc=ed,dc=ac,dc=uk"
        $samAccountName = "$pcName-rdp"
    }

    # Try and create the group
    try {
        Write-Host "`nAttempting to create $groupType group for $pcName"
        New-ADGroup -Name $pcName -SamAccountName  $samAccountName -GroupCategory Security -GroupScope Global -Path $OU -Description $description -ErrorAction SilentlyContinue
        Write-host "`n$groupType group for $pcName created at the following location :"
        Write-Host "$OU`n" -ForegroundColor Green
    }
    # If unable to create the group then display error message
    catch {
        Write-Host "`nUnable to create group!" -ForegroundColor Red
        Write-Host "Cause :$_.`n" -ForegroundColor Red
    }

}

function Create-macOSAdminGroup {
    param([Parameter(Mandatory=$true)][string]$PCName)

    # For Macs, we already know the group type is admin
    $groupType = "Administrators"

    # Set samAccountName and sub OU
    $samAccountName = $PCName
    $subOU = "LF"

    # Set OU Path
    $OU = "OU=$groupType,OU=$subOU,OU=Authorisation,OU=UoESD,DC=ed,dc=ac,dc=uk"
    # Ask for description
    $description = Read-Host "`nPlease enter a description for the group - normally building, user and [P] for permanent "

    # Try and create the group
    try {
        Write-Host "`nAttempting to create $groupType group...."
        New-ADGroup -Name $pcName -SamAccountName  $samAccountName -GroupCategory Security -GroupScope Global -Path $OU -Description $description -ErrorAction SilentlyContinue
        Write-host "`n$groupType group for $pcName created at the following location :" -ForegroundColor Green
        Write-Host "$OU`n" -ForegroundColor Green
        return "success"
    }
    # If unable to create the group then display error message
    catch {
        Write-Host "`nUnable to create group!" -ForegroundColor Red
        Write-Host "Cause :$_.`n" -ForegroundColor Red
        return "error"
    }
}

