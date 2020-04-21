function Display-VirtualizationInBIOSHP {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    # First, Enter-PSSession on remote machine        
    # Create new PS Session
    $Session = New-PSSession $PCName

    #Invoke command on remote machine
    Invoke-Command -Session $Session -ScriptBlock {        

        # Get BIOS Settings
        $bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biosEnumeration -EnableAll

        # Get individual settings
        $VTx = ($bios |?{$_.Name -like "Virtualization Technology (VTx)"}).Value
        $Vtd = ($bios |?{$_.Name -like "Virtualization Technology Directed I/O (VTd)"}).Value

        # Display current state of virtualiaztion settings'
        Write-Host "`nCurrent Virtualization settings on $Using:PCName"
        Write-Host "-------------------------------------------------"
        Write-Host "Virtualization Technology (VTx) : $VTx" 
        Write-Host "Virtualization Technology Directed I/O (VTd) : $VTd" 
        Write-Host ""
    } 

    # Remove PS Session
    Remove-PSSession $Session

}

function Enable-VirtualizationInBIOSHP {
    param ([Parameter(Mandatory=$true)][string]$PCName)

     # Check PC name is valid
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check if PC is online. If not then return
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }    
    
    do {
        $BIOSPword = Check-BIOSPword
    } until ($BIOSPword -ne "error")  

    # So far so good, so lets attempt the change! 
    # Create new PS Session    
    $Session = New-PSSession $PCName    

    #Invoke command on remote machine

    Invoke-Command -Session $Session -ScriptBlock {

        # Nested function to display settings
        function Display-Settings{
            # Get BIOS Settings
            $bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biosEnumeration -EnableAll

            # Get individual settings
            $VTx = ($bios |?{$_.Name -like "Virtualization Technology (VTx)"}).Value
            $VTd = ($bios |?{$_.Name -like "Virtualization Technology Directed I/O (VTd)"}).Value

            # Display current state of virtualiaztion settings'
            Write-Host "`nVirtualization settings on $Using:PCName"
            Write-Host "-------------------------------------------"
            Write-Host "Virtualization Technology (VTx) : $VTx" 
            Write-Host "Virtualization Technology Directed I/O (VTd) : $VTd" 
            Write-Host ""
        }
        
        # Display current settings before change
        Write-Host "`nObtaining current settings..."
        (Display-Settings $Using:PCName)

        # Create interface to change BIOS Settings
        $bios2 = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface -EnableAll

        Write-Host "Performing changes..."

        # Enable BIOS Setting that matches name Virtualization Technology (VTx)
        $VTxExecuteChangeAction = $bios2.SetBIOSSetting('Virtualization Technology (VTx)', 'Enable', $Using:BIOSPWord)

        # Get return code. If it's not 0 then there's been a problem
        $VTxReturnCode = $VTxExecuteChangeAction.return

        if(($VTxReturnCode) -eq 0){
            Write-Host "`nVTx setting appears to have successfully changed!" -ForegroundColor Green
        }
        else {
            Write-Host "`nUnable to change VTx setting. You can attempt it manually through a PowerShell Session`n" -ForegroundColor Red
        }

        # Enable BIOS Setting that matches name Virtualization Technology Directed I/O (VTd)
        $VTdExecuteChangeAction = $bios2.SetBIOSSetting('Virtualization Technology Directed I/O (VTd)', 'Enable', $Using:BIOSPWord)

        # Get return code. If it's not 0 then there's been a problem
        $VTdReturnCode = $VTdExecuteChangeAction.return

        if(($VTdReturnCode) -eq 0){
            Write-Host "`nVTd setting appears to have successfully changed!" -ForegroundColor Green
        }
        else {
            Write-Host "`nUnable to change VTd setting. You can attempt it manually through a Remote PowerShell Session`n" -ForegroundColor Red

        }

        if (($VTxReturnCode -ne 0) -or ($VTdReturnCode) -ne 0){
            Write-Host "`nThere was a problem changing either one or both of the settings. You may have to do this manually.`n" -ForegroundColor Red
            return
        }
        else {
            Write-Host "`nDone! Re-displaying settings to check if they have changed..." -ForegroundColor Green

            # Display settings again after change
            (Display-Settings $Using:PCName)
        }       
    }
    
    # Remove PS Session
    Remove-PSSession $Session

}

function Disable-VirtualizationInBIOSHP {
    param ([Parameter(Mandatory=$true)][string]$PCName)
    
    # Check PC name is valid
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check if PC is online. If not then return
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }    
    
    do {
        $BIOSPword = Check-BIOSPword
    } until ($BIOSPword -ne "error")    

    # First, Enter-PSSession on remote machine        
    # Create new PS Session
    $Session = New-PSSession $PCName

    #Invoke command on remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        
        # Nested function to display settings
        function Display-Settings{
            # Get BIOS Settings
            $bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biosEnumeration -EnableAll

            # Get individual settings
            $VTx = ($bios |?{$_.Name -like "Virtualization Technology (VTx)"}).Value
            $VTd = ($bios |?{$_.Name -like "Virtualization Technology Directed I/O (VTd)"}).Value

            # Display current state of virtualiaztion settings'
            Write-Host "`nVirtualization settings on $Using:PCName"
            Write-Host "-------------------------------------------------"
            Write-Host "Virtualization Technology (VTx) : $VTx" 
            Write-Host "Virtualization Technology Directed I/O (VTd) : $VTd" 
            Write-Host ""
        }

        

        # Display current settings before change
        Write-Host "`nObtaining current settings..."
        (Display-Settings $Using:PCName)

        # Create interface to change BIOS Settings
        $bios2 = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface -EnableAll

        Write-Host "Performing changes..."

        # Disable BIOS Setting that matches name Virtualization Technology (VTx)
        $VTxExecuteChangeAction = $bios2.SetBIOSSetting('Virtualization Technology (VTx)', 'Disable', $Using:BIOSPWord)

        # Get return code. If it's not 0 then there's been a problem
        $VTxReturnCode = $VTxExecuteChangeAction.return

        if(($VTxReturnCode) -eq 0){
            Write-Host "`nVTx setting appears to have successfully changed!" -ForegroundColor Green
        }
        else {
            Write-Host "`nUnable to change VTx setting. You can attempt it manually through a PowerShell Session`n" -ForegroundColor Red
        }

        # Disable BIOS Setting that matches name Virtualization Technology Directed I/O (VTd)
        $VTdExecuteChangeAction = $bios2.SetBIOSSetting('Virtualization Technology Directed I/O (VTd)', 'Disable', $Using:BIOSPWord)

        # Get return code. If it's not 0 then there's been a problem
        $VTdReturnCode = $VTdExecuteChangeAction.return
    
        if(($VTdReturnCode) -eq 0){
            Write-Host "`nVTd setting appears to have successfully changed!" -ForegroundColor Green
        }
        else {
            Write-Host "`nUnable to change VTd setting. You can attempt it manually through a PowerShell Session`n" -ForegroundColor Red

        }

        if (($VTxReturnCode -ne 0) -or ($VTdReturnCode) -ne 0){
            Write-Host "`nThere was a problem changing either one or both of the settings. You may have to do this manually.`n" -ForegroundColor Red
            return
        }
        else {
            Write-Host "`nDone! Re-displaying settings to check if they have changed..." -ForegroundColor Green

            # Display settings again after change
            (Display-Settings $Using:PCName)
        }

    }

    # Remove PS Session
    Remove-PSSession $Session

}

function Disable-BIOSLockHP {
    param ([Parameter(Mandatory=$true)][string]$PCName)
    # Show disclaimer
    Write-Host "`nPlease be aware that BIOS lock is not available for any computers previous to G2 models, so this will not work on G1 or prior models.`n" -ForegroundColor

     # Check PC name is valid
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check if PC is online. If not then return
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }

    # Check 
    
    # Loop until correct BIOS password is supplied or quit
    do {
        $BIOSPword = Check-BIOSPword
    } until($BIOSPword -ne "error")

    # Get all BIOS settings
    $bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biosEnumeration -EnableAll

    # Display status of BIOS lock
    Write-Host "`nCurrent BIOS lock status on $PCName : " ($bios |?{$_.Name -like "Lock BIOS Version"}).Value
    Write-Host ""

    $bios2 = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface -EnableAll

    $bios2.SetBIOSSetting('Lock BIOS Version', 'Disable', $BIOSPword)

    $bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biosEnumeration -EnableAll

    Write-Host "`nNew BIOS lock status on $PCName : " ($bios |?{$_.Name -like "Lock BIOS Version"}).Value
    Write-Host ""
}

function Check-BIOSPword{
    #$KFile = "\\cmvm.datastore.ed.ac.uk\cmvm\mvmsan\med-apps\ISLFSupport\PowerShell\k\K.key"
    $PFile = "\\cmvm.datastore.ed.ac.uk\cmvm\mvmsan\med-apps\ISLFSupport\PowerShell\p\P.txt"
    $Key = New-Object Byte[] 16
    $BIOSPWord_ = Get-Content $PFile | ConvertTo-SecureString -Key $Key

    # Prompt for BIOS password and store in $Password as a secure string
    $Password_ = Read-Host "`nEnter BIOS password (or ctrl+c to quit)" -AsSecureString

    # convert secure string to normal string
    $BIOSPWord = "<utf-16/>"+[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($BIOSPWord_))
    $Password = "<utf-16/>"+[Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password_))

    # Make sure passwords match otherwise password is incorrect
    if ($Password -cne $BIOSPWord){
        Write-Host "`nIncorrect BIOS Password!" -ForegroundColor Red
        return "error"
    }
    else{
        return $BIOSPWord    
    }
}

function Disable-BIOSPasswordHP {
    

}