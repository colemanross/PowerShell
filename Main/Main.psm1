function Add-AdminRights {
    
    # Set flag to catch errors
    $FAILED = "N"

    Write-Host "`nADD ADMIN RIGHTS"
    Write-Host "=================="
    # Get PC name. Loop until valid name supplied.
    do {
        $PCName_ = Read-Host "Please enter the PC Name (or CTRL+C to exit)"
        $PCName = $PCName_.ToUpper()
        $PCExists = (Check-ADComputer $PCName)
        if ($PCExists -eq "error") {
            Write-Host "Cannot find $PCName in Active Directory. Please enter a valid PC name!`n" -ForegroundColor Red
        }
    } until ($PCExists -ne "error")

    # Check if AD admin group already exists
    $groupExists = (Check-ADGroup $PCName)
    # If not, do we want to create one? Loop until valid answer is given.
    if ($groupExists -eq "error") {
        do {
            Write-Host "There appears to be no administrator group for PC $PCName in AD." -ForegroundColor Yellow
            $question = "`nDo you wish to create one (y/n) ?"
            $confirm = (Confirm-Answer $question)
        } until ($confirm -ne "error")
        # If not, quit
        if ($confirm -eq "y"){
            (Create-Win10Group $PCName "Administrators")
        }
    }
    # Else group already exists
    else {
        Write-Host "`nAdmin group for $PCName exists." -ForegroundColor Green
    }

    # Add the uun to the group
    $uun = (Add-ADUserToADGroup $PCName "administrator")

    # Ask if changes should be forced down to local remote machine
    do {
        $force = "Do you wish to attempt to force rights down to the local PC (y/n) ?" 
        $confirmForce = (Confirm-Answer $force)
    } until ($confirmForce -ne "error")
    # If not then quit
    if ($confirmForce -eq "y") {
        $PCOnline = (Check-PCisOnline $PCName)
        # If not then attempt to wake.
        if ($PCOnline -eq "error") {
            Write-Host "$PCName appears to be offline. Attempting to wake. This may take around 30 seconds...." -ForegroundColor Yellow
            (Wake-RemotePC $PCName)
            Start-Sleep 30
            $PCOnline = (Check-PCisOnline $PCName)
            if ($PCOnline -eq "error") {
                Write-Host "`nSorry, this process is unable to wake $PCName or it is taking longer then 30 secs to wake." -ForegroundColor Red
            }
        }
        # Else, attempt to obtain the local admin group
        else {
            try {    
                $localGroup = [ADSI]"WinNT://$PCName/Administrators,group" 
            }
            catch {
                Write-Host "`nUnable to obtain information from $PCName." -ForegroundColor Red
                Write-Host "Cause $_.`n" -ForegroundColor Red
                return
            }


            # Get current members of the local administrators group
            $tempMembers = @($localGroup.psbase.Invoke("Members"))
            # Store each member of the local administrators group 
            $members = $tempMembers | foreach {$_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null)}
            # Check if the admin group we are changing is already there. If so, attempt to remove so we can force new changes in the AD group down. 
            if ($PCName -in $members) {
                Write-Host "AD admin group $PCName already exists in local admin group on $PCName."
                Write-Host "Attempting to remove and then re-add to make sure changes take place..."
                try {
                    ([ADSI]"WinNT://$PCName/Administrators,group").remove("WinNT://ed.ac.uk/$PCName") # Remove group 
                }
                catch {
                    Write-Host "`nUnable to remove $PCName from local $groupName group on PC $PCName!" -ForegroundColor Red
                    Write-Host "Cause :" $_. -ForegroundColor Red
                    $FAILED = "Y"
                }                
            }
            # Sleep for a few seconds just to make sure change takes place
            Start-Sleep 3
            # Attempt to add the admin AD group to the local admin group
            try {
                $localGroup.Add("WinNT://ed.ac.uk/$PCName,group") #Add the AD Group to the local group      
                Write-Host "`nAD group $PCName has been successfully added to the local admin group on PC $PCName`n" -ForegroundColor Green # Display successfull message 
            }
            catch {
                Write-Host "`nUnable to add $PCName to Administrators group on PC $PCName!" -ForegroundColor Red
                Write-Host "Cause : $_." -ForegroundColor Red
                $FAILED = "Y"
            }
        }
        
    }

    # If for some reason there was an issue forcing rights locally, ask if Local Users and Groups should be opened on the rmeote machine
    if ($FAILED -eq "Y") {
        # loop until valid answer is given
        do {
            Write-Host "There appeared to be an issue forcing rights down to the local machine" -ForegroundColor Yellow
            $openGroups = "`nDo you wish to open the Local Users and Groups service on $PCName to check local permissions (y/n) ?"
            $confirmOpenGroup = (Confirm-Answer $openGroups) 
        } until ($confirmOpenGroup -ne "error")
        if ($confirmOpenGroup -eq "y"){
            # Open Local Users and Groups for remote machine
            lusrmgr.msc /computer="$PCName"
        }
    }

    # Ask if standard reply should be copied to clipboard
    do {
        $firstName = (Return-FirstName $uun)
        $standardReply = "Dear $firstName,

We have now granted you temporary admin rights on $PCName. You will need to restart the computer for them to take effect.

Please let us know once you have finished with the rights."
    
        $copyStandardReply = "`nDo you wish to copy a standard reply to your clipboard (y/n) ?"
        $confirmReply = (Confirm-Answer $copyStandardReply)
    } until ($confirmReply -ne "error")

    if ($confirmReply -eq "y") {
        $standardReply | clip.exe
    } 
}

function Remove-AdminRights {
    Write-Host "`nREMOVE ADMIN RIGHTS"
    Write-Host "===================="
    # Get PC name. Loop until valid name supplied.
    do {
        $PCName = Read-Host "Please enter the PC Name (or CTRL+C to exit)"
        $PCExists = (Check-ADComputer $PCName)
        if ($PCExists -eq "error") {
            Write-Host "Cannot find $PCName in Active Directory. Please enter a valid PC name!`n" -ForegroundColor Red
        }
    } until ($PCExists -ne "error")

    # Check if AD admin group already exists
    $groupExists = (Check-ADGroup $PCName)
    if ($groupExists -eq "error") {
        Write-Host "`nThere does not appear to be an existing administrator group in AD for $PCName!`n" -ForegroundColor Red
    }
    else{
        (Display-ADGroupMembers $PCName)
        (Remove-ADUserFromADGroup $PCName)
    }
}

function Add-RemoteDesktopUser {

    # Set flag to catch errors
    $FAILED = "N"

    Write-Host "`nADD REMOTE DESKTOP USER"
    Write-Host "---------------------------"
    # Get PC name. Loop until valid name supplied.
    do {
        $PCName_ = Read-Host "Please enter the PC Name (or CTRL+C to exit)"
        $PCName = $PCName_.ToUpper()
        $PCExists = (Check-ADComputer $PCName)
        if ($PCExists -eq "error") {
            Write-Host "Cannot find $PCName in Active Directory. Please enter a valid PC name!`n" -ForegroundColor Red
        }
    } until ($PCExists -ne "error")

    # Check if AD RDP group already exists
    $groupExists = (Check-ADGroup "$PCName-rdp")
    # If not, do we want to create one? Loop until valid answer is given.
    if ($groupExists -eq "error") {
        do {
            Write-Host "There appears to be no Remote Desktop Users group for PC $PCName in AD." -ForegroundColor Yellow
            $question = "`nDo you wish to create one (y/n) ?"
            $confirm = (Confirm-Answer $question)
        } until ($confirm -ne "error")
        # If so, create group
        if ($confirm -eq "y"){
            (Create-Win10Group $PCName "RemoteUsers")
        }
    }
    # Else group already exists
    else {
        Write-Host "`nRemote Desktop User group for $PCName exists." -ForegroundColor Green
    }

    # Add the uun to the group
    $uun = (Add-ADUserToADGroup "$PCName-rdp" "Remote Desktop Users")

    # Ask if changes should be forced down to local remote machine
    do {
        $force = "Do you wish to attempt to force rights down to the local PC (y/n) ?" 
        $confirmForce = (Confirm-Answer $force)
    } until ($confirmForce -ne "error")
    # If so then attempt it.
    if ($confirmForce -eq "y") {
        # If so, check PC is online
        Write-Host "This can take a few seconds. Please be patient."  
        $PCOnline = (Check-PCisOnline $PCName)
        # If not then try to wake.
        if ($PCOnline -eq "error") {
            Write-Host "$PCName appears to be offline. Attempting to wake. This may take around 30 seconds...." -ForegroundColor Yellow
            (Wake-RemotePC $PCName)
            Start-Sleep 30
            $PCOnline = (Check-PCisOnline $PCName)
            if ($PCOnline -eq "error") {
                Write-Host "`nSorry, this process is unable to wake $PCName or it is taking longer then 30 secs to wake." -ForegroundColor Red
            }
        }
        # Else, attempt to obtain the local rdp group
        else {
            try {    
                $localGroup = [ADSI]"WinNT://$PCName/Remote Desktop Users,group" 
            }
            catch {
                Write-Host "`nUnable to obtain information from $PCName." -ForegroundColor Red
                Write-Host "Cause $_.`n" -ForegroundColor Red
                return
            }

            # Get current members of the local rdp group
            $tempMembers = @($localGroup.psbase.Invoke("Members"))
            # Store each member of the local rdp group 
            $members = $tempMembers | foreach {$_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null)}
            # Check if the rdp group we are changing is already there. If so, attempt to remove so we can force new changes in the AD group down. 
            if ($PCName -in $members) {
                Write-Host "AD Remote Desktop Users group $PCName already exists in local Remote Desktop Users group on $PCName."
                Write-Host "Attempting to remove and then re-add to make sure changes take place..."
                try {
                    ([ADSI]"WinNT://$PCName/Remote Desktop Users,group").remove("WinNT://ed.ac.uk/$PCName") # Remove group 
                }
                catch {
                    Write-Host "`nUnable to remove $PCName from local $groupName group on PC $PCName!" -ForegroundColor Red
                    Write-Host "Cause :" $_. -ForegroundColor Red
                    $FAILED = "Y"
                }                
            }
            # sleep for a few seconds to make sure change takes place
            Start-Sleep 5
            # Attempt to add the RDP AD group to the local admin group
            try {
                $localGroup.Add("WinNT://ed.ac.uk/$PCName-rdp,group") #Add the AD Group to the local group      
                Write-Host "`nAD group $PCName-rdp has been successfully added to the local Remote Desktop Users group on PC $PCName`n" -ForegroundColor Green # Display successfull message 
            }
            catch {
                Write-Host "`nUnable to add $PCName to Remote Desktop Users group on PC $PCName!" -ForegroundColor Red
                Write-Host "Cause : $_." -ForegroundColor Red
                $FAILED = "Y"
            }
        } 
    }
    
     # If for some reason there was an issue forcing rights locally, ask if Local Users and Groups should be opened on the rmeote machine
    if ($FAILED -eq "Y") {
        # loop until valid answer is given
        do {
            Write-Host "There appeared to be an issue forcing rights down to the local machine" -ForegroundColor Yellow
            $openGroups = "`nDo you wish to open the Local Users and Groups service on $PCName to check local permissions (y/n) ?"
            $confirmOpenGroup = (Confirm-Answer $openGroups) 
        } until ($confirmOpenGroup -ne "error")
        if ($confirmOpenGroup -eq "y"){
            # Open Local Users and Groups for remote machine
            lusrmgr.msc /computer="$PCName"
        }
    } 
    
    # Ask if RDP file should be created. Loop until valid answer.
    do {
        $RDPFile = "Do you wish to create an .RDP file for $uun on $PCName (y/n) ?"
        $confirmRDPFile = (Confirm-Answer $RDPFile)
    } until ($confirmRDPFile -ne "error")

    if ($confirmRDPFile -eq "y") {
        do {
            $OSVersion = "Is this for an off-campus Windows computer (y/n) ? (n = macOS)"
            $confirmOS = (Confirm-Answer $OSVersion)
        } until ($confirmOS -ne "error")

        # If 'y', then create windows file
        if ($confirmOS -eq "y") {
            (Create-RDP $PCName $uun)
        }
        # If n then create mac files
        if ($confirmOS -eq "n") {
        (Create-macOSRDP $PCName $uun)
        }
    }

    # Populate standard reply. Loop until valid answer is given.
    do {
        # Get forename
        $firstName = (Return-FirstName $uun)
        # Populate standard reply
        $standardPCReply = "Dear $firstName,

We have now enabled your remote desktop access to computer $PCName. We've also attached a small .zip file (note - this will only work on a Windows computer). On your off-campus computer, if you download, unzip and run the file it will then attempt to connect to your office computer. You will need to make sure that your office computer is live online, so you should make sure that you add the device to your MyEd wake list to be able to wake the device from a remote location :

http://www.docs.is.ed.ac.uk/docs/Subjects/IS-Help/wake-on-lan.pdf

Please note that you can only add your office computer to your MyEd wake list from the office computer itself. Once added you can then login to MyEd from your off-campus machine and wake your office PC.

Please let us know if it works."
    
        $copyStandardReply = "`nDo you wish to copy a standard reply to your clipboard (y/n) ?"
        $confirmReply = (Confirm-Answer $copyStandardReply)
    } until ($confirmReply -ne "error")

    # If 'y', then copy reply to clipboard.
    if ($confirmReply -eq "y") {
        $standardPCReply | clip.exe
    }      
}

function Add-macOSAdminRights {
    
    Write-Host "`nAdd macOS Admin rights"
    Write-Host "=================="

    do{
        $desktopQuestion = "Is this for a macOS Desktop (y/n) (n = laptop) ?"
        $confirmDesktop = (Confirm-Answer $desktopQuestion)
    } until ($confirmDesktop -ne "error")

    # If it's a desktop make sure AD Computer object exists
    If ($confirmDesktop -eq "y"){
        # Get PC name. Loop until valid name supplied.
        do {
            $PCName_ = Read-Host "Please enter the PC Name (or CTRL+C to exit)"
            $PCName = $PCName_.ToUpper()
            $PCExists = (Check-ADComputer $PCName)
            if ($PCExists -eq "error") {
                Write-Host "Cannot find $PCName in Active Directory. Please enter a valid PC name!`n" -ForegroundColor Red
            }
        } until ($PCExists -ne "error")
    }

    # If it's a laptop, make sure it has typical naming convention
    If ($confirmDesktop -eq "n") {
        # Create flag to make sure laptop has correct school code
        $FOUND = "N"
        do {
            $PCName_ = Read-Host "Please enter the macOS laptop Name (or CTRL+C to exit)"
            $PCName = $PCName_.ToUpper()
            $PCNameLength = $PCName.Length
            $maximumCharacters = 12
            if ($PCNameLength -ne $maximumCharacters){
                Write-Host "`nInvalid character length for a macOS laptop name! There should be 12 characters, you have entered a name with $PCNameLength characters." -ForegroundColor Red
            }
        } until ($PCName.Length -eq $maximumCharacters)

        $schoolCodes = @('S31','S32','S33','S34','S35','S37','P5L')
        foreach ($code in $schoolCodes) {
            if ($PCName.StartsWith($code)) {
                $FOUND = "Y"
                break
            } 
        }
        If ($FOUND -eq "N"){
            Write-Host "`nThe name you have entered does not follow the normal naming convention for macOS laptop devices.`n" -ForegroundColor Red
            Write-Host "The name should begin with one of the following School codes :" -ForegroundColor Yellow
            foreach ($code in $schoolCodes){
                Write-Host "`t$code" -ForegroundColor Yellow
            }
            Write-Host "`nThis process will now exit`n" -ForegroundColor Red
            return 
        }
    }

    # Check if AD admin group already exists
    $groupExists = (Check-ADGroup $PCName)
    # If not, do we want to create one? Loop until valid answer is given.
    if ($groupExists -eq "error") {
        do {
            Write-Host "There appears to be no administrator group for macOS device $PCName in AD." -ForegroundColor Yellow
            $question = "`nDo you wish to create one (y/n) ?"
            $confirm = (Confirm-Answer $question)
        } until ($confirm -ne "error")
        # If not, quit
        if ($confirm -eq "n"){
            return    
        }
        # Else, attempt to create group
        else {
            $success = (Create-macOSAdminGroup $PCName)
            if ($success -eq "error"){
                Write-Host "`nUnfortunately this process has been unable to create admin group $PCname" -ForegroundColor Red
                Write-Host "You will need to create the group manually."
                Write-Host "This process will now exit.`n"
                return
            }
        }        
    }
    # Else group already exists
    else {
        # Get OU path
        $OUPath = (Return-GroupOUPath $PCName)
        do {
            Write-Host "`nAdmin group for $PCName exists in the following location :"
            Write-Host "`n$OUPath"
            $question = "`nDo you wish to use this group (y/n) ?"
            $confirm = (Confirm-Answer $question)
        } until ($confirm -ne "error")

        if ($confirm -eq "n") {
            return
        }
    }

    # Add the uun to the group
    $uun = (Add-ADUserToADGroup $PCName "administrator")
         
    
    
}