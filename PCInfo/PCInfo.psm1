# Displays general information about the PC
function Display-PCInfo {
    <#
        .Synopsis
        Displays information about a remote PC.

        .Description
        Displays information about a remote PC, such as Manufacturer, Model, serial number etc..
        Also display information on Operating System, Processor, RAM, BIOS, Hard Drive usage and if any user is currently logged in.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>
    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Nested function
    function Display-Info {
        param ([Parameter(Mandatory=$true)][string]$PCName)
        # Set Error preference
        $ErrorActionPreference = "Stop"
        try {
            # Display Information
            "COMPUTER NAME : $PCName"
            ""
            "Product Details"
            "======================================================"
            "Manufacturer : " + (Get-WmiObject Win32_ComputerSystem -ComputerName $PCName).Manufacturer # Obtain the Computer Manufacturer
            "Model : " + (Get-WmiObject Win32_ComputerSystem -ComputerName $PCName).Model # Obtain the Computer Model
            "Serial Number / Service Tag (DELL): " + (Get-WmiObject Win32_bios -ComputerName $PCName).SerialNumber # Obtain serial number
            "Product Number (HP) : " + (Get-WmiObject -Computername $PCName -Namespace Root\wmi MS_SystemInformation).SystemSKU #
            ""
            "Operating System Details"
            "======================================================"
            "OS Version : " + (Get-WmiObject Win32_OperatingSystem -ComputerName $PCName).Caption # Display Operating System
            "OS Architecture : " + (Get-WmiObject Win32_OperatingSystem -ComputerName $PCName).OSArchitecture # Display OS Architecture
            ""   
            "Processor Details"
            "======================================================"
            "Processor Information : " + (Get-WmiObject Win32_Processor -ComputerName $PCName).Name # Display Processor information
            "Processor Architecture : " + (Get-WmiObject Win32_Processor -ComputerName $PCName).AddressWidth + "-bit" # Display Processor Architecture
            ""
            "RAM"
            "======================================================"
            $RAM = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb
            "Approximate RAM : $RAM GB" # Display approximate RAM
            ""
            "BIOS Information"
            "======================================================"
            "BIOS Manufacturer : " + (Get-WmiObject Win32_bios -ComputerName $PCName).Manufacturer # Display BIOS Manufacturer
            "BIOS Version : " + (Get-WmiObject Win32_bios -ComputerName $PCName).BIOSVersion # Display BIOS Version
            "BIOS SMBIOSBIOSVersion : " + (Get-WmiObject Win32_bios -ComputerName $PCName).SMBIOSBIOSVersion # Display SMBIOS version
            ""
            "Hard Drive Information"
            "======================================================"
            $disk = ([wmi]"\\$PCName\root\cimv2:Win32_logicalDisk.DeviceID='c:'")
            "The C: drive/partition on $PCName has approximately {0:#.0} GB free of {1:#.0} GB Total" -f ($disk.FreeSpace/1GB),($disk.Size/1GB) # Display hard drive info
            ""   
            "Current Logged in User(s)"
            "====================================================="
            "Logged on console user(s) : " + (Get-WmiObject Win32_ComputerSystem  -ComputerName $PCName).Username # Display currently logged in username
            ""
        }
        catch {
            Write-Host "Unable to display $PCName details!" -ForegroundColor Red
            Write-Host "Cause : $_.`n" -ForegroundColor Red
        }
    }

    # ===== Begin main function =====

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check if PC is online. If not then return
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    # Else PC is online. Let's get the info!
    else {
        # If no file currently exists in C:\Workspace\$pcName, go ahead and display PC info
        if (-Not (Test-Path C:\Workspace\$PCName.txt)){
            (Display-Info $PCName | Tee-Object -file C:\Workspace\$PCName.txt)
            Invoke-Item C:\Workspace\$PCName.txt
            Get-Content C:\Workspace\$PCName.txt | clip
            Write-Host "`nOutput has been saved to C:\Workspace\$pcName.txt and content copied to clip board" -ForegroundColor Green 
        }
        # Else it appears that there is aleready a file named C:\Workspace\$PCName. Ask to overwrite.
        else {
            do {
                Write-Host "`nC:\Workspace\$PCName.txt already exists. Running this will overwrite this file." -ForegroundColor Yellow
                $overwrite = "`nDo you wish to continue ( y/n ) (select n to display to console only) ? "
                $confirm = (Confirm-Answer $overwrite)
            } until ($confirm -ne "error")
        } 
        # Switch statement to control output of computer info
        switch($confirm) {
            "y" { 
                try {
                    (Display-Info $PCName | Tee-Object -file C:\Workspace\$PCName.txt)
                    Invoke-Item C:\Workspace\$PCName.txt
                    Get-Content C:\Workspace\$PCName.txt | clip
                    Write-Host "Output has been saved to C:\Workspace\$pcName.txt and content copied to clip board" -ForegroundColor Green
                }
                catch {
                    Write-Host "`nUnable to output to file!" -ForegroundColor Red
                    Write-Host "Cause : $_." -ForegroundColor Red   
                }
            }
            "n" {(Display-Info $PCName)}
        }
    }   
}

function Display-SerialInfo {

    <#
        .Synopsis
        Displays Serial information for a remote PC.

        .Description
        Displays Serial information for a remote PC. For Dells, it will show Service Tags. For HPs it will show Product number.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        # Set Error action preference
        $ErrorActionPreference = "Stop"

        # Try and obtain the serial details
        try {
            Write-Host "`nSerial for $PCName"
            Write-Host "======================================================"
            Write-Host "Serial Number / Service Tag (DELL): " (Get-WmiObject Win32_bios -ComputerName $PCName).SerialNumber # Obtain serial number
            Write-Host "Product Number (HP) : " (Get-WmiObject -Computername $PCName -Namespace Root\wmi MS_SystemInformation).SystemSKU #        
        }
        catch {
            Write-Host "`nUnable to obtain serial!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
        }
    }
}

function Return-HPSerial {

    <#
        .Synopsis
        Returns serial number for a remote HP PC.

        .Description
        Returns serial number for a remote HP PC.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        $ErrorActionPreference = "Stop"
        try {
            $HPSerial = (Get-WmiObject Win32_bios -ComputerName $PCName).SerialNumber # Obtain serial number
            return $HPSerial                
        }
        catch {
            Write-Host "`nUnable to obtain serial!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    }
}

function Return-HPProductNumber {

    <#
        .Synopsis
        Returns Product number for a remote HP PC.

        .Description
        Returns Product number for a remote HP PC.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        $ErrorActionPreference = "Stop"
        try {        
            $productNumber = (Get-WmiObject -Computername $PCName -Namespace Root\wmi MS_SystemInformation).SystemSKU #
            return $productNumber       
        }
        catch {
            Write-Host "`nUnable to obtain Product Number for $PCName!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    }     
}

function Return-DellSerialInfo {

    <#
        .Synopsis
        Returns serial number for a remote Dell PC.

        .Description
        Returns serial number for a remote Dell PC.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }
    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        $ErrorActionPreference = "Stop"
        try {
            $dellSerial = (Get-WmiObject Win32_bios -ComputerName $PCName).SerialNumber # Obtain serial number
            return $dellSerial                
        }
        catch {
            Write-Host "`nUnable to obtain serial for $PCName!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    }
}

function Display-OSVersion {
    
    <#
        .Synopsis
        Displays Operating System version for a remote PC.

        .Description
        Displays Operating System version for a remote PC.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }    
    else {
        $ErrorActionPreference = "Stop"
        try {
            Write-Host "`nOperating System Details"
            Write-Host "======================================================"
            Write-Host "OS Version : "(Get-WmiObject Win32_OperatingSystem -ComputerName $PCName).Caption # Display Operating System
            Write-Host "OS Architecture : "(Get-WmiObject Win32_OperatingSystem -ComputerName $PCName).OSArchitecture # Display OS Architecture        
        }
        catch {
            Write-Host "`nUnable to obtain serial!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    }
}

function Return-OSVersion {
    
    <#
        .Synopsis
        Returns Operating System version for a remote PC.

        .Description
        Returns Operating System version for a remote PC.
                
        REQUIREMENTS:
        - Remote PC must be online
        - PC Object must exist in AD
        - You must have admin rights on remote machine 

    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        $ErrorActionPreference = "Stop"
        try {
            $OSVersion = (Get-WmiObject Win32_OperatingSystem -ComputerName $PCName).Caption
            return $OSVersion               
        }
        catch {
            Write-Host "`nUnable to OS Version from $PCName!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    }
}

function Display-MACAddressFromAD{
    
    <#
        .Synopsis
        Displays mac address attribute for an AD PC.

        .Description
        Displays mac address attribute for an AD PC Object.
                
        REQUIREMENTS:
        - PC Object must exist in AD
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }
    else {   
        $ErrorActionPreference = "Stop"
        try {
            # Query mac address attribute
            $MACAddress = (Get-ADComputer $PCName -property extensionAttribute2).ExtensionAttribute2
            # Display Mac address
            Write-Host "`nMAC Address of AD PC $PCName is $MACAddress and has been copied to your clipboard`n"
            # Copy Mac address to clip board
            $MACAddress | clip.exe                        
        }
        catch {
            Write-Host "`nUnable to obtain MAC address for $PCName from AD object!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    }
}

function Return-MACAddressFromAD {
    
    <#
        .Synopsis
        Returns mac address attribute for an AD PC.

        .Description
        Returns mac address attribute for an AD PC Object.
                
        REQUIREMENTS:
        - PC Object must exist in AD
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }
    else {
        $ErrorActionPreference = "Stop"
        try {
            $MACAddress = (Get-ADComputer $PCName -property extensionAttribute2).ExtensionAttribute2
            return $MACAddress                      
        }
        catch {
            Write-Host "`nUnable to obtain MAC address for $PCName from AD object!" -ForegroundColor Red
            Write-Host "Cause :$_.`n" -ForegroundColor Red
            return
        }
    } 
}

function Display-EdLANDBDetails {}

function Display-ARP {}

function Display-MACAddressInEdLANDB {
    
}

function Return-MACAddressInEdLANDB {
    
}

function Display-InstalledApps {

    <#
        .Synopsis
        Displays installed Apps on a remote PC.

        .Description
        Displays installed Apps on a remote PC. Outputs in Grid view for 32-bit, 64-bit and AppV installs.
                
        REQUIREMENTS:
        - PC Object must exist in AD.
        - Remote PC must be online.       
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)
    
    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }    
    # Else PC is online
    else {
        Write-Host "`n$PCName still appears to be online. Attempting to obtain installed applications..." -ForegroundColor Green
        # Invoke commmand allows us to run script blocks on remote computers
        try{
            # Display 32-bit Applications
            Invoke-Command -computer $PCName {Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object DisplayName} | Out-GridView -Title "32-bit Applications installed on $pcName"
        }
        catch {
            Write-Host "`nUnable to obtain 32-bit Apps on $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }
        try {
            # Display 64-bit Applications
            Invoke-Command -computer $PCName {Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object DisplayName} | Out-GridView -Title "64-bit Applications installed on $pcName"
        }
        catch {
            Write-Host "`nUnable to obtain 64-bit Apps on $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }
        try {
            # Display Virtual Apps (*** NOT SURE IF THIS WORKS ***)
            Invoke-Command -computer $PCName {Get-AppvClientApplication | Select-Object Name, Version | Sort-Object Name} | Out-GridView -Title "Virtual Applications installed on $pcName"
        }
        catch {
            Write-Host "`nUnable to obtain Virtual Apps on $PCName!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }
    }
    Write-Host "Done!" -ForegroundColor Green
}

function Display-LastBootTime {
    
    <#
        .Synopsis
        Displays last boot time of a remote PC.

        .Description
        Displays last boot time of a remote PC.
                
        REQUIREMENTS:
        - PC Object must exist in AD
        - Remote PC must be online
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }
    else {
        $PCOnline = (Check-PCIsOnline $PCName)
        if ($PCOnline -eq "error") {
            Write-Host "`n$PCName is offline. Exiting..." -ForegroundColor Red
            return
        }
        else {
            try {
                Get-CimInstance -ClassName win32_operatingsystem -computername $PCName | select csname, lastbootuptime
            }
            catch {
                Write-Host "`nUnable to obtain last boot time!" -ForegroundColor Red
                Write-Host "Cause : $_.`n" -ForegroundColor Red
            }
        }
    }

}

function Display-LastLoggedInUser {
    
    <#
        .Synopsis
        Displays last logged in user of a remote PC.

        .Description
        Displays last logged in user of a remote PC. 
        Note that all it does is look for the last modified NTUSER.DAT file in C:\Users, so may not always display correctly the last logged in user.
                
        REQUIREMENTS:
        - PC Object must exist in AD
        - Remote PC must be online
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        $uun = (Get-Childitem \\$PCName\c$\users\*\ntuser.dat -Force |?{$_.Directory -notmatch "(Default|Public|TestUser)"}|Sort LastWriteTime | Select -Last 1).Directory|Select Name, LastWriteTime
        $username = ($uun.Name).Trim()
        $fullName = (Return-FullName $username)
        $result = New-Object PSObject -Property @{Name=$PCName; UUN = $uun.Name ; FullName = $fullName ; LastWriteTime = $uun.LastWriteTime}
        $results += $result
        Write-Host "`nLast NTUSER.DAT file to be modified :"
        $results
    }   
}

function Display-LoggedInUsers {

    <#
        .Synopsis
        Displays all logged in users of a remote PC.

        .Description
        Displays all logged in users of a remote PC. Looks to see if explorer.exe is running and if so displays the owner of the process.
                
        REQUIREMENTS:
        - PC Object must exist in AD
        - Remote PC must be online
    #>

    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        # Get explorer processes
        $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ComputerName $PCName -ErrorAction SilentlyContinue)
        # If there are no explorer processes found then no one appears to be logged in
        if ($explorerprocesses.Count -eq 0) {
            Write-Host "No explorer process found / Nobody interactively logged on"
        }
        # Else, for each explorer process obtain the owner and (most likely!) the time they logged in
        else {
            foreach ($i in $explorerprocesses) {
                $Username = $i.GetOwner().User
                $Domain = $i.GetOwner().Domain
                Write-Host "`n$Domain\$Username logged on since"($i.ConvertToDateTime($i.CreationDate))  -ForegroundColor Yellow
            }
        }
    }

}

function Display-GraphicsAdapter{
    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $PCName)
    if ($PCExists -eq "error") {
        return
    }

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        Get-WmiObject -ComputerName "$PCName" win32_VideoController
    }
        
}