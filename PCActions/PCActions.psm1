function Rename-ComputerName{

    <#
        .Synopsis
        Renames Computer

        .Description
        Remotely renames a computer which is connected to the university domain.
        
        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        .Example
        - Obtains uun of user calling the function
        - Asks for current PC name to rename
        - Checks to make sure PC exists in AD
        - Checks to make sure PC is online
        - Requests the new name
        - Checks to make sure new name doesn't currently exists in AD
        - Renames the computer

    #>

    # Display requirements
    Write-Host "****************************************************" -ForegroundColor Yellow
    Write-Host "** Before running this script, please make sure   **" -ForegroundColor Yellow
    Write-Host "** that changes have been made in EdLAN DB, and   **" -ForegroundColor Yellow
    Write-Host "** that the current computer object is located    **" -ForegroundColor Yellow
    Write-Host "** in the correct Organisational Unit in Active   **" -ForegroundColor Yellow
    Write-Host "** Directory.                                     **" -ForegroundColor Yellow
    Write-Host "**                                                **" -ForegroundColor Yellow
    Write-Host "** Do NOT change the name of the object in Active **" -ForegroundColor Yellow
    Write-Host "** Directory. This script will perform this.      **" -ForegroundColor Yellow
    Write-Host "****************************************************" -ForegroundColor Yellow

    Write-Host "`n** This script will also automatically restart the computer **`n" -ForegroundColor Yellow

    # Get current logged in user who will be executing the name change
    $loggedInUser = whoami

    do {
        $currentPCName_=read-host "Please enter CURRENT computer name (or enter q to quit) " # Get the old PC name which is to get changed
        $currentPCName = $currentPCName_.ToUpper()
        If($currentPCName -eq "Q"){ # If entry equals "q" then exit program
            return
        }
        $validatePCName = (Check-ADComputer $currentPCName)
        if ($validatePCName -eq "error") {            
            Write-Host "`n$currentPCName does not exist in Active Directory. Please enter a valid Computer name.`n" -ForegroundColor Red        
        }
    } until ($validatePCName -eq "PC exists")

    # Check PC exists in AD. If not return.
    $PCExists = (Check-ADComputer $currentPCName)
    if ($PCExists -eq "error") {
        return
    }
    # Check PC is online
    $PCOnline = (Check-PCIsOnline $currentPCName)
    if ($PCOnline -eq "error"){
        return
    }

    # Get new PC details
    do {
        $newPCName_ = Read-Host "Please enter NEW PC name (or enter q to quit) "
        $newPCName = $newPCName_.ToUpper()
        If($newPCName -eq "Q"){ # If entry equals "q" then exit program
            return
        }

        $validatePCName = (Check-ADComputer $newPCName)
        if ($validatePCName -eq "error") {        
            Write-Host "Computer name $newPCName is available" -ForegroundColor Green        
        }
        else {        
            Write-Host "There is already a computer in Active Directory called $newPCName. Please enter a different name." -ForegroundColor Red        
        }
    } until ($validatePCName -eq "error")

    do {
        $question = "Are you sure you wish to change $currentPCName to $newPCName (y / n) ? "
        $confirm = (Confirm-Answer $question)
    } until ($confirm -ne "error")

    if ($confirm -eq "n"){
        return
    }
    else {
        # Get the logged in user
        $loggedInUser = whoami
        # Get Credentials
        $myCredential = Get-Credential $loggedinuser
        try {
            Write-Host "Attempting to change name..."
            Rename-Computer -ComputerName $currentPCName -NewName $newPCName -DomainCredential $myCredential -Force -PassThru -Restart -ErrorAction Stop
            Write-Host "*** Computer name change appears to have succeeded. The remote computer should now be restarting.***" -ForegroundColor Green
        }
        catch {
            Write-Host $_.
            Write-Host "`n*** ERROR *** Unable to change name. This is usually due to one of the following reasons :" -ForegroundColor Red
            Write-Host "`t 1. If attempting to rename remotely, the PC is now offline" -ForegroundColor Red
            Write-Host "`t 2. The new PC name now has an object created and is therefore now unavailable to you" -ForegroundColor Red
            Write-Host "`t 3. If attempting to rename remotely, you do not have admin priveledges either on the remote computer or on the domain`n" -ForegroundColor Red
        }
    }

}

function Wake-RemotePC { 
        param ([Parameter(Mandatory=$true)][string]$PCName)
        # Get IP Address
        $IPAddress = [System.Net.Dns]::GetHostAddresses("$PCName").IPAddressToString
        try {
            start-process -wait "http://wol.is.ed.ac.uk/wol/index.cgi?IPADDRESS=$IPAddress" -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "`nCannot contact WOL service. Unable to contact url :`n" -ForegroundColor Red
            Write-Host "`thttp://wol.is.ed.ac.uk/wol/" -ForegroundColor Yellow
            Write-Host "`nAre you certain you are online and have authorisation to use the service?`n" -ForegroundColor Red
        }
}

function Copy-FullData {

    <#
        .Synopsis
        Copies all data within a specific folder.

        .Description
        Asks for a logfile name and output location.
        Copies all data within a specific folder to a location of your choice. Includes any subfolders. No files or folder are excluded.
                
        REQUIREMENTS:
        - Any remote PCs / Servers need to be online.

        .Example
        - Asks for a locaiton to save the log file.
        - Asks to input name for the copy log.
        - Asks for source folder location
        - Asks for target location
        - Preforms Robocopy

    #>

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select location to save log file"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK"){
        $logLocation += $foldername.SelectedPath
    }
    else {
        return
    }    

    $logName = Read-Host "Enter name for log file "
    $log = "$logLocation\$logName.log"
    Write-Host "`nLog file will be created at $log" -ForegroundColor Yellow
    $sourcePath = Read-Host "`nSource folder (eg C:\Workspace\Folder 1) "
    $destinationPath = Read-Host "Destination folder (eg C:\Users\user\Local Documents\Folder 1) - You will need to enter a folder name to be created  "
    ROBOCOPY "$sourcePath" "$destinationPath" /E /XJ /COPY:DAT /LOG:$log /V /ETA /TEE /R:1 /W:5
}

function Copy-DataExcludingOutlook {

    <#
        .Synopsis
        Copies data within a specific folder but excludes .ost and .pst file extensions (Outlook).

        .Description
        Asks for a logfile name and output location.
        Copies all data within a specific folder to a location of your choice. Includes any subfiles and subfolders. Excludes any .ost and .pst files (Outlook) .
                
        REQUIREMENTS:
        - Any remote PCs / Servers need to be online.

        .Example
        - Asks for a locaiton to save the log file.
        - Asks to input name for the copy log.
        - Asks for source folder location
        - Asks for target location
        - Preforms Robocopy

    #>

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select location to save log file"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK"){
        $logLocation += $foldername.SelectedPath
    }
    else {
        return
    }    

    $logName = Read-Host "Enter name for log file "
    $log = "$logLocation\$logName.log"
    Write-Host "`nLog file will be created at $log" -ForegroundColor Yellow
    $sourcePath = Read-Host "`nSource folder (eg C:\Workspace\Folder 1) "
    $destinationPath = Read-Host "Destination folder (eg C:\Users\user\Local Documents\Folder 1) - You will need to enter a folder name to be created  "
    ROBOCOPY "$sourcePath" "$destinationPath" /E /XJ /XF *.ost *.pst /COPY:DAT /LOG:$log /V /ETA /TEE /R:1 /W:5
}

function Copy-DataResume {
    
    <#
        .Synopsis
        Resumes a copy job that has been interrupted.

        .Description
        Asks for a logfile name and output location.
        Resumes a copy job that has been interrupted, as long as correct paths have been supplied.
                
        REQUIREMENTS:
        - Any remote PCs / Servers need to be online.

        .Example
        - Asks for a locaiton to save the log file.
        - Asks to input name for the copy log.
        - Asks for source folder location
        - Asks for target location
        - Preforms Robocopy

    #>

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select location to save log file"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK"){
        $logLocation += $foldername.SelectedPath
    }
    else {
        return
    }    

    $logName = Read-Host "Enter name for log file "
    $log = "$logLocation\$logName.log"
    Write-Host "`nLog file will be created at $log" -ForegroundColor Yellow
    $sourcePath = Read-Host "`nSource folder (eg C:\Workspace\Folder 1) "
    $destinationPath = Read-Host "Destination folder (eg C:\Users\user\Local Documents\Folder 1) - You will need to enter a folder name to be created  "
    ROBOCOPY "$sourcePath" "$destinationPath" /E /XJ /Z /COPY:DAT /LOG:$log /V /ETA /TEE /R:1 /W:5
}

function Create-RDP {
    <#
        .Synopsis
        Creates a pre-configured RDP file for a PC.

        .Description
        Creates a pre-configured RDP file for a PC, including details of remote gateway and username.
                
        REQUIREMENTS:
        - N/A

        .Paramters
        NetBIOS name required.

        .Example
        - Checks to make sure PC exists in AD
        - Asks for uun of user requesting remote access
        - Checks to make sure uun is valid in AD
        - Creates .rdp file with uun and gateway settings
        - Asks for location to save .rdp file

    #>
    param (
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$false)][string]$uun
    )

    # We need to use a form for Save dialog, so import the library
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    # Set error perference
    $ErrorActionPreference = 'SilentlyContinue'
    
    # If we are running the Create-RDP command from the PS prompt, we need to check machine name and username is valid. These checks should have been performed if being called from another function.
    if ($uun -eq "") {
        # Check PC exists in AD. If not return.
        $PCExists = (Check-ADComputer $PCName)
        if ($PCExists -eq "error") {        
            return
        }
        
        $uun = Read-Host "`nPlease enter uun "
        $userExists = (Check-ADuser $uun)
        If ($userExists -eq "error") {
            return
        }
    }

        # Populate file with defaults and entries
        $defaultText = "screen mode id:i:1
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1157
session bpp:i:16
winposstr:s:0,3,480,0,1280,600
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:2
displayconnectionbar:i:1
disable wallpaper:i:1
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:$PCName
audiomode:i:0
redirectprinters:i:1
redirectcomports:i:0
redirectsmartcards:i:1
redirectclipboard:i:1
redirectposdevices:i:0
redirectdirectx:i:1
autoreconnection enabled:i:1
authentication level:i:0
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:rd-gateway.is.ed.ac.uk
gatewayusagemethod:i:1
gatewaycredentialssource:i:0
gatewayprofileusagemethod:i:1
promptcredentialonce:i:1
use redirection server name:i:0
networkautodetect:i:1
bandwidthautodetect:i:1
enableworkspacereconnect:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
gatewaybrokeringtype:i:0
drivestoredirect:s:*
username:s:ed\$uun"

        # We want to remove the dns extension for the name of the file (it's neater :) ). Use "." as a separator so we can delete everything after it
        $separator = "."
        $separatePCName = $PCName.split($separator)
        # Get up to first "." and convert to upper case
        $netBIOS = $separatePCName[0].ToUpper()

        # Create SaveAsFile Dialog, with default .rdp extension and filename. Default save location is C:\Workspace
        $SaveChooser = New-Object -Typename System.Windows.Forms.SaveFileDialog
        $SaveChooser.Filter = "RDP File|*.rdp"
        $SaveChooser.AddExtension = ".rdp"
        $SaveChooser.FileName = $netBIOS
        $SaveChooser.InitialDirectory = "C:\Workspace"
        # Show dialog
        $SaveChooser.ShowDialog()
        # Write content and create file
        $defaultText | Set-Content $SaveChooser.FileName
        
        # Try and zip
        $filePath = $SaveChooser.FileName
        Compress-Archive -Path $filePath -DestinationPath "$filePath.zip"
        
        # Now that we have it zipped, delete original
        Remove-Item $filePath
        
      
}

# Function to create .rdp file
function Create-macOSRDP {
    param (
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$false)][string]$uun
    )

    # We need to use a form for Save dialog, so import the library
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    # Create an internal function to write out .rdp preferences. This enables us to create more than one .rdp file.
    function Create-RDPPrefs {

        param (
            [Parameter(Mandatory=$true)][string]$PCName,
            [Parameter(Mandatory=$false)][string]$uun,
            [Parameter(Mandatory=$true)][string]$gateway
        )

        $defaultText = "screen mode id:i:1
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1157
session bpp:i:16
winposstr:s:0,3,480,0,1280,600
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:2
displayconnectionbar:i:1
disable wallpaper:i:1
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:$PCName
audiomode:i:0
redirectprinters:i:1
redirectcomports:i:0
redirectsmartcards:i:1
redirectclipboard:i:1
redirectposdevices:i:0
redirectdirectx:i:1
autoreconnection enabled:i:1
authentication level:i:0
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:$gateway
gatewayusagemethod:i:1
gatewaycredentialssource:i:0
gatewayprofileusagemethod:i:1
promptcredentialonce:i:1
use redirection server name:i:0
networkautodetect:i:1
bandwidthautodetect:i:1
enableworkspacereconnect:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
gatewaybrokeringtype:i:0
drivestoredirect:s:*
username:s:ed\$uun"

return $defaultText
        
    }

    # Set error perference
    $ErrorActionPreference = 'SilentlyContinue'
    
    # If we are running the Create-RDP command from the PS prompt, we need to check machine name and username is valid. These checks should have been performed if being called from another function.
    if ($uun -eq "") {
        # Check PC exists in AD. If not return.
        $PCExists = (Check-ADComputer $PCName)
        if ($PCExists -eq "error") {        
            return
        }
        
        $uun = Read-Host "`nPlease enter uun "
        $userExists = (Check-ADuser $uun)
        If ($userExists -eq "error") {
            return
        }
    }
    
    # Declare array of different gateways
    $gateways = @("toran.is.ed.ac.uk","portico.is.ed.ac.uk","ianua.is.ed.ac.uk","vrata.is.ed.ac.uk","doras.is.ed.ac.uk","gonhi.is.ed.ac.uk","rd-gateway.is.ed.ac.uk")
    

        # Declare target path for temporary folder. This will be used to store .rdp files and then copmressed.
        $targetPath_ = "C:\Windows\Temp\"
        # Get timestamp
        $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
        # Create temporary folder with timestamp as name so we are sure it's a unique name
        $targetPath = New-Item -Path $targetPath_ -Name "$PCName-$timestamp" -ItemType "directory"
        
        # We want to remove the dns extension for the name of the file (it's neater :) ). Use "." as a separator so we can delete everything after it
        $separator = "."
        $separatePCName = $PCName.split($separator)
        # Get up to first "." and convert to upper case
        $netBIOS = $separatePCName[0].ToUpper()

        $number = 1
        foreach ($gateway in $gateways) {
            $file = (Create-RDPPrefs $netBIOS $uun $gateway)            
            $fullPath = New-Item -Path $targetPath -Name "$netBIOS-$number.rdp" -ItemType file
            Set-Content -Path $fullPath -Value $file
            $number++    
        }

        $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
        $foldername.Description = "Select location to save .zip file"
        $foldername.rootfolder = "MyComputer"

        if($foldername.ShowDialog() -eq "OK"){
            $saveLocation += $foldername.SelectedPath
        }
        else {
            return
        }

        # Rename the folder
        Rename-Item $targetPath $netBIOS

        # Get new name
        $newTargetPath = "C:\Windows\Temp\$netBIOS"
        
        # Compress and send to save location
        Compress-Archive -Path $newTargetPath -DestinationPath "$saveLocation\$netBIOS.zip"
        
        # Now that we have it zipped, delete original
        Remove-Item $newTargetPath -Recurse
      
}

function Restart-RemoteComputer {
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

    # Check to make sure no one is logged in
    # Declare array to store logged in usernames
    $loggedInusers = @()
    
    # Get windows explorer processes from remote computer
    $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ComputerName $PCName -ErrorAction SilentlyContinue)
    # If no explorer processes are running then most likely no one is logged in
    if ($explorerprocesses.Count -eq 0) {
        Write-Host "Currently, no one appears to be logged into $PCName."
        do {
            $question = "`nDo you wish to continue with a restart (y/n) ?"
            $answer = (Confirm-Answer $question)
        } until ($answer -ne "error")

        if ($answer -eq "n"){
            return
        }
    }
    else {
        # Else for each explorer process that is running, get the username of the owner of the process
        foreach ($i in $explorerprocesses) {
            # Get the username
            $Username = $i.GetOwner().User
            # Get the domain
            $Domain = $i.GetOwner().Domain
            $fullUsername = "$Domain\$Username"
            $loggedInusers += $fullUsername           
        }
        # Display logged in users
        $displayArray = $loggedinUsers -join ', '
        Write-Host "`nThe following users have explorer.exe process running on $PCName, which indicates that they are currently logged in :"
        Write-Host "`n$displayArray`n"
        Write-Host "*** WARNING *** - Any unsaved data will be lost!" -ForegroundColor Yellow
        # Loop until valid reply
        do {
            $question = "`nDo you wish to continue (y/n) ?"
            $answer = (Confirm-Answer $question)
        } until ($answer -ne "error")

        if ($answer -eq "n"){
            return
        }
    }

    # Perform restart
    Restart-Computer -ComputerName $PCName -force
}


function Remove-VirtualApp {
    
}

function Remove-VirtualAppsAll {    

    Write-Host "`nThe following script will remove all installed App-V applications on the remote computer.`n" -ForegroundColor Yellow

    $PCName = Read-Host "Enter computer name on which to remove Virtual Apps "

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

    # Get the logged in user
    $loggedInUser = whoami
    # Get Credentials
    $myCredential = Get-Credential $loggedinuser

    # Create new PS Session
    $Session = New-PSSession $PCName -Credential $myCredential
   
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {

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
               
        # Get all installed App-V Apps
        $VApps = Get-AppvClientPackage

        # If there are no AppV apps then continue
        if ($VApps.Count -eq 0) {
            Write-Host "`nNo Virtual Apps appear to be installed.`n" -ForegroundColor Green
        }
        # Else, we need to warn the user that these apps will be removed
        else {

            Write-Host "The following Virtual Applications appear to be installed :`n" -ForegroundColor Yellow
            $i = 1
            ForEach ($app in $VApps){
                Write-Host `t $i - $app.Name
                $i += 1               
            }

            Write-Host ""
            do {
                $question = "If you continue then these applications will be removed. Do you wish to continue (y/n) ? "
                $confirm = (Confirm-Answer $question)
            } until ($confirm -ne "error")
            if ($confirm -eq "n"){
                return
            }
            else {
                # Unpublish all apps
                Write-Host "Unpublishing all App-V Apps..."
                Unpublish-AppvClientPackage *
                Write-Host "Removing all App-V Apps..."
                Remove-AppvClientPackage *
                Write-Host "Done."                             
            }
        }


    }
    
    # Remove PS Session
    Remove-PSSession $Session
     
}

function Remove-AppVClient {
    

    Write-Host "`nThe following script will remove the App-V Client on the remote computer.`n"
    Write-Host "Subsequently, this also remove any install Virtual Apps!`n" -ForegroundColor Yellow
    Write-Host "Please also note that this processwill restart the computer! Any unsaved data will be lost!`n" -ForegroundColor Yellow

    $PCName = Read-Host "Enter computer name on which to remove Virtual Apps "

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

    $Session = New-PSSession $PCName

    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        
        # Get location of AppvClient
        $AppVClient = ""

        Write-Host "`nUn-installing App-V Client on $Using:PCName...."  
        Start-Process -wait -FilePath $AppVClient -ArgumentList "/UNINSTALL" 
        Write-Host "`nApp-V Client un-installed. Attempting to restart $Using:PCName...."  
    }
    
    # Remove PS Session
    Remove-PSSession $Session 
    
    (Restart-RemoteComputer $PCName)   
}

function Remove-CCMClient {
    Write-Host "This script will remove the CCM client on a remote computer."
    Write-Host "`nPlease note that this process will restart the copmuter! Any unsaved data will be lost!`n" -ForegroundColor Yellow

    # Get PC name
    $PCName = Read-Host "Enter computer name on which to remove the CCM Client "

    # Declare a location for the log
    #$logFile = "C:\Workspace\$PCName-CCM-Remove.log"

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

    $Session = New-PSSession $PCName

    

    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $logFile)

        # Nested function for deleting folders
        function Remove-Folder {
            param ([Parameter(Mandatory=$true)][string]$Path)

            # Check folder exists
            If (Test-Path -Path $Path){
                Write-Host "`nRemoving $Path..."
                try {
                    Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "`nUnable to remove $Path! You may need to remove manually!" -ForegroundColor Red
                    Write-Host "Cause : $_." -ForegroundColor Red
                }
            }
            else {
                Write-Host "`n$Path doesn't appear to exist!" -ForegroundColor Yellow
            }        
        }

        # Nested function for removing Registry key        
        function Remove-RegistryKey{
            param ([Parameter(Mandatory=$true)][string]$regKey)

            # Find key
            Write-Host "`nAttempting to find key...`n" 
            if (Get-Item -Path $regKey -ErrorAction SilentlyContinue){
                Write-Host "$regKey exists. Removing...`n" -ForegroundColor Green
                try { 
                    Remove-Item -Path $regKey -Recurse -ErrorAction SilentlyContinue
                } catch {
                    Write-Host "`nUnable to remove $regKey!" -ForegroundColor Red
                    Write-Host "Cause : $_." -ForegroundColor Red
                }
            }
            else {
                Write-Host "$regKey does not exist.`n" -ForegroundColor Red
                
            }              
        }

        # Nested function to remove Scheduled Task
        function Remove-ScheduledTask {
            param ([Parameter(Mandatory=$true)][string]$STName)

            # Create Task Scheduler COM object
            $ST = New-Object -ComObject Schedule.Service
            # Gonnect to local task sceduler 
            $ST.Connect($env:COMPUTERNAME) 
            # Get tasks folders 
            $RootFolder = $ST.GetFolder("\")
            $CMTaskFolder = $ST.GetFolder("\Microsoft\Configuration Manager")           
            
            # If supplied argument is all, then remove all Configuration Manager tasks, along with folder itself
            if ($STName -eq "all") {
                Write-Host "`nRemoving all scheduled tasks in \Microsoft\Configuration Manager\" 
                $CMTasks = $CMTaskFolder.GetTasks(0)
                Foreach($Task in $CMTasks) {
                        $CMTaskFolder.DeleteTask($Task.Name,$null) 
                }
                # Remove the folder
                Write-Host "`nRemoving \Microsoft\Configuration Manager\ folder..."
                $RootFolder.DeleteFolder("Microsoft\Configuration Manager",$null)
            }
            # Else we are removing individual tasks
            else {
                Write-Host "Removing scheduled task $STName..."
                $RootTasks = $RootFolder.GetTasks(0)
                Foreach ($Task in $RootTasks) {
                    if ($Task.Name -eq $STName) {
                        $RootFolder.DeleteTask($Task.Name,$null)
                    }
                }
            }

        }

        # Check file - If this file no longer exists then CCM Client has most likely successfully uninstalled. 
        $checkfile = "c:\windows\CCM\Logs\PolicyEvaluator.Log"
        
        # Declare location of CCM client
        $CCMClient = "C:\Windows\CCMSetup\ccmsetup.exe"

        # Uninstall Client
        Write-Host "`nUninstalling CCM Client..." 
        # Start-Process -wait -FilePath $CCMClient -ArgumentList "/uninstall"
        & $CCMClient /uninstall | Out-Null

        # Loop until checkfile doesn't exist or tries = 100 (5 minutes)
        $looptries = 1
        do{
		    Start-Sleep -Seconds 3
		    $looptries++
		} until (!(Test-Path -Path $checkfile) -or ($looptries -eq $100))
       
        # Stop the SMS Agent Host service
        Write-host "`nStopping SMS Agent Host service..."
        try {
            Stop-Service CcmExec -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "`nUnable to stop SMS Agent Host Service!" -ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }

        # Removing folders
        $CCM = "C:\Windows\CCM"
        (Remove-Folder $CCM)

        $CCMSetup = "C:\Windows\ccmsetup"
        (Remove-Folder $CCMSetup)
        
        $CCMCache = "C:\Windows\ccmcache"
        (Remove-Folder $CCMCache)

        $smscfg = "C:\Windows\SMSCFG.INI"
        (Remove-Folder $smscfg)

        $smsmif = "C:\Windows\SMS*.mif"
        (Remove-Folder $smsmif)      
        
        # Remove registry keys
        Write-Host "`nRemoving Registry Keys..." 

        $CCMKey = 'HKLM:\SOFTWARE\Microsoft\CCM'
        (Remove-RegistryKey $CCMKey)

        $CCMSetupKey = 'HKLM:\SOFTWARE\Microsoft\CCMSetup'
        (Remove-RegistryKey $CCMSetupKey)

        $SMSSetupKey = 'HKLM:\SOFTWARE\Microsoft\SMS'
        (Remove-RegistryKey $SMSSetupKey)

        # Remove Namespaces
        #SMS
        $smsNameSpace = (Get-WmiObject -query "Select * From __Namespace Where Name='sms'" -Namespace "root\cimv2").Name
        if ($smsNameSpace -eq "") {
            Write-Host "`nroot\cimv2\sms doesn't appear to exist."
        }
        else {
            Write-Host "`nRemoving root\cimv2\sms Namespace..." 
            Get-WmiObject -query "Select * From __Namespace Where Name='sms'" -Namespace "root\cimv2" | Remove-WmiObject -verbose
        }
        # CCM
        $ccmNameSpace = (Get-WmiObject -query "Select * From __Namespace Where Name='ccm'" -Namespace "root").Name
        if ($ccmNameSpace -eq "") {
            Write-Host "`nroot\ccm doesn't appear to exist." 
        }
        else {
            Write-Host "`nRemoving root\ccm Namespace..." 
            Get-WmiObject -query "Select * From __Namespace Where Name='ccm'" -Namespace "root" | Remove-WmiObject -verbose
        }

        # Removed Configuration Manager Scheduled Tasks
        (Remove-ScheduledTask "all")

        # Remove individual tasks
        (Remove-ScheduledTask "Windows10MigrationNotice")

        # Remove Certificates
        # Get-ChildItem -Path Cert:\LocalMachine\SMS\ | Remove-Item

         
    } -ArgumentList $PCName, $logFile

    # Remove PS Session
    Remove-PSSession $Session

    # Restarting
    Write-Host "`nRestarting..."
    (Restart-RemoteComputer $PCName)

    # Loop until offline so we can make sure machine restarts. Give it a timeout of 10 minutes (updates may need to install).
    #$offlineCounter = 1
    #do{
    #    $ping = Test-Connection -ComputerName $PCName -Quiet
    #    $counter++      
    #} until (!$ping -or $offlineCounter -eq 600)

    #if ($offlineCounter -eq 600){
    #   Write-Host "`n$PCName doesn't appear to have shutdown within 10 minutes! This script will quit. You will need to remove manually.`n" -ForegroundColor Red | Tee-Object -FilePath $logFile
    #    return
    #}    
    #Write-Host "`n$PCName successfully powered down. Waiting for device to come back online..." | Tee-Object -FilePath $logFile

    # Next loop until machine back online. Give it a timeout of 10 minutes.
    #$onlineCounter = 1     
    # do{
    #    $ping = Test-Connection -ComputerName $PCName -Quiet         
    #} until ($ping -or $onlineCounter -eq 600)

    #if ($onlineCounter -eq 600) {
    #    Write-Host "`n$PCName has taken over 10 minutes to come back online, indicating an issue! This script will quit. You will need to remove manually.`n" -ForegroundColor Red | Tee-Object -FilePath $logFile
    #    return
    #}   

}