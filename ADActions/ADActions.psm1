function Add-ToMDSDGroup {
    <#
        .Synopsis
        Adds a PC AD Object to an MVM Mobile Desktop group.

        .Description
        Adds a PC AD Object to an MVM Mobile Desktop group. Mainly required for laptop builds to gain "Direct Access".

        REQUIREMENTS:
        - Write Permissions to various MDSD AD Groups.

        WHAT IT DOES:
        - Determines if OS Version is 10 or another (7 or 8.1).
        - If not Win 10, asks what MDSD Group to add the PC to.
        - Checks to make sure PC Object doesn't currently exist in the MDSD group.
        - Adds the PC Object to the selected MDSD group.
        - If Win 10, checks to make sure PC Object doesn't already exist.
        - Adds PC Object directly to "ed.ac.uk/UoEX/Auth/MVM/MVM SDX Mobile Desktop".
    #>
      
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$OSVersion
    )

# Nested function for adding computer to group
    function addToGroup([string]$PCName,[string]$groupName, [array]$currentMembers) {
        # For each member of the group
        foreach ($pc in $currentMembers){
            # Check if the PC name supplied already exists in the group
            if ($pc -eq $PCName){
                Write-Host "`n$PCName already appears to be in $groupName!`n" -ForegroundColor Yellow
                return
            }
        }  
        # Lets add it!  
        $PCObject = Get-ADComputer $PCName
        try {
            Add-ADGroupMember $groupName -Members $PCObject
            Write-Host "Sucessfully added $PCName to $groupName." -ForegroundColor Green
        }
        catch {
            Write-Host "`nUnable to add $PCName to $groupName!" -ForegroundColor Red
            Write-Host "Cause : $_.`n"
        }                  
    }

    # If OS Version is not equal to Windows 10
    if ($OSVersion -ne "Microsoft Windows 10 Education"){
        # Ask for which MDSD group to add the machine into. Loop until a valid answer is supplied
        do {
            Write-Host "`nWhat MDSD Client group do you wish to add $pcName to? "
            Write-Host "`n`t 1. ISD MDSD Clients"
            Write-Host "`t 2. MVMCOLLEGE MDSD Clients"
            Write-Host "`t 3. CCBS MDSD Clients"
            Write-Host "`t 4. CPHS MDSD Clients"
            Write-Host "`t 5. SBMS MDSD Clients"
            Write-Host "`t 6. BRR MDSD Clients"
            Write-Host "`t 7. CRFR MDSD Clients"
            Write-Host "`t 8. CRIC MDSD Clients"
            $answer_ = Read-Host "`nEnter option (q to go back) "
            # Convert to upper case
            $answer = $answer_.ToUpper()
            # Set value of the MDSD group depending on answer
            switch ($answer){
                1 {$MDSDGroupName = "ISD MDSD Clients"}
                2 {$MDSDGroupName = "MVMCOLLEGE MDSD Clients"}
                3 {$MDSDGroupName = "CCBS MDSD Clients"}
                4 {$MDSDGroupName = "CPHS MDSD Clients"}
                5 {$MDSDGroupName = "SBMS MDSD Clients"}
                6 {$MDSDGroupName = "BRR MDSD Clients"}
                7 {$MDSDGroupName = "CRFR MDSD Clients"}
                8 {$MDSDGroupName = "CRIC MDSD Clients"}
                # If user selects Q, then break from switch statement
                "Q" {break}
                # If user enters anything other than 1 -> 9 or q, then display error
                default {Write-Host "`n*** Invalid entry! Please enter a number from 1 - 8 (or q to go back). ***`n" -ForegroundColor Red}
            }
            # If user selects Q, exit from this script
            if ($answer -eq "Q") {exit}
        } until (($answer -in 1 .. 8) -or ($answer -eq "Q"))

        #Create array to hold members of selected MDSD group
        [array]$MDSDCurrentMembers = (Get-ADGroupMember $MDSDGroupName).Name
        # Add the machine to the MDSD group by calling the function
        (addToGroup $pcName $MDSDGroupName $MDSDCurrentMembers)    
    }

    # If Machine is Windows 10, then no need to select a specific MDSD group. Add to the default SDX mobile group
    if ($OSVersion -eq "Microsoft Windows 10 Education"){
        [array]$MVMCurrentMembers = (Get-ADGroupMember "MVM SDX Mobile Desktop").Name
        (addToGroup $pcName "MVM SDX Mobile Desktop" $MVMCurrentMembers)
    }   
}

function Disable-ADComputer {
    <#
        .Synopsis
        Disables an AD PC Object.

        .Description
        Disables an AD PC Object.

        REQUIREMENTS:
        - AD edit permissions on the PC object being disabled.

        WHAT IT DOES:
        - Obtains timestamp
        - Obtains uun from logged in user calling function
        - Checks to make sure PC Object exists in AD
        - Attempts to disable object

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Get time stamp and logged in user
    $date = $((Get-Date).ToString())
    $uun = whoami

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }    
    else {
        $fullName = "$PCName$"
        try{
            Disable-ADAccount -Identity $fullName
            Write-Host "$PCName sucessfully disabled!" -ForegroundColor Green
        }
        catch {
            Write-Host "Can't disable $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }
        try{
            $description = "# DISABLED by $uun on $date"
            Set-ADComputer $PCName -Description $description -ErrorAction Stop
        }
        catch {
            Write-Host "`nUnable to set description for $PCName`n" -ForegroundColor Red
            Write-Host "Cause : $_.`n" -ForegroundColor Red
        }
    }
}

function Disable-ADComputers {
    <#
        .Synopsis
        Disables multiple PC Objects.

        .Description
        Disables multiple PC Objects. Reads PC names from a .txt file. Machines must be on separate lines in the .txt file.

        REQUIREMENTS:
        - .txt file containing list of machines to be disabled.
        - AD edit permissions on the PC objects being disabled.

        WHAT IT DOES
        - Asks for location of .txt file
        - Iterates through list one by one disabling objects.
        - Also fills out description field with user name and timestamp of when the object was disabled
    #>

    Write-Host "`nPlease make sure that your imported file is .txt format, and that each machine name is on a separate line.`n" -ForegroundColor Yellow

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'Text (*.txt)|*.txt'
    }
    $null = $FileBrowser.ShowDialog()

    $hosts = Get-Content $FileBrowser.FileName

    $date = $((Get-Date).ToString())
    $uun = whoami

    foreach ($comp in $hosts) {
        $fullName = $comp+'$'
        try{
            Disable-ADAccount -Identity $fullName
            Write-Host "`n$comp disabled successfully`n" -ForegroundColor Green
        }
        catch{
            Write-Host "`n$_.`n" -ForegroundColor Red
        }
        try{
            $description = "# DISABLED by $uun on $date`n"
            Set-ADComputer $comp -Description $description -ErrorAction Stop
        }
        catch {
            Write-Host "Unable to set description for $comp`n" -ForegroundColor Red
        }
    }    
}

function Enable-ADComputer{

    <#
        .Synopsis
        Enables an AD PC Object.

        .Description
        Enables an AD PC Object.

        REQUIREMENTS:
        - AD edit permissions on the PC object being enabled.

        WHAT IT DOES:
        - Obtains timestamp
        - Obtains uun from logged in user calling function
        - Checks to make sure PC Object exists in AD
        - Attempts to enable object

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Get time stamp and logged in user
    $date = $((Get-Date).ToString())
    $uun = whoami
    
    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }    
    else {
        $fullName = "$PCName$"
        try{
            Enable-ADAccount -Identity "$fullName"
            Write-Host "$PCName sucessfully enabled!" -ForegroundColor Green
        }
        catch {
            Write-Host "`nCan't enable $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_.`n" -ForegroundColor Red
        }
        try{
            $description = "# ENABLED by $uun on $date"
            Set-ADComputer $PCName -Description $description -ErrorAction Stop
        }
        catch {
            Write-Host "`nUnable to set description for $PCName" -ForegroundColor Red
            Write-Host "Cause : $_.`n" -ForegroundColor Red
        }
    }

}

function Display-Win7BLRecoveryKeyPassword{
    Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase "OU=Laptops,OU=LF,OU=MVM,OU=ISD,OU=SD7,OU=UoESD,DC=ed,DC=ac,DC=uk" -Properties 'msFVE-RecoveryPassword'| Where-Object {$_.DistinguishedName –match "8C6CCE96-E0BB-4009-AA0B-05639B3CBA6A"}
}

function Display-Win7BLRecoveryKeyName{

    param ([Parameter(Mandatory=$true)][string]$PCName)

   Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase "OU=Laptops,OU=LF,OU=MVM,OU=ISD,OU=SD7,OU=UoESD,DC=ed,DC=ac,DC=uk" -Properties 'msFVE-RecoveryPassword, whenCreated,whenChanged'| Where-Object {$_.DistinguishedName –match $PCName}
}

function Display-BLRecoveryKeyName{

    param ([Parameter(Mandatory=$true)][string]$PCName)

   Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase "OU=BQ,OU=MVM,OU=ISSD,OU=SDX,OU=UoEX,DC=ed,DC=ac,DC=uk" -Properties 'msFVE-RecoveryPassword, whenCreated,whenChanged'| Where-Object {$_.DistinguishedName –match $PCName}
}