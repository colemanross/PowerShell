function Display-Scans {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Display paths
    get-childitem -recurse \\ed\dst\logs\ | ?{$_.name -match $PCName} | select fullname
    Write-Host "`n"
    # Obtain paths
    $fullPath = (get-childitem -recurse \\ed\dst\logs\ | ?{$_.name -match $PCName}).FullName
    # For each item, dispay the details
    foreach ($item in $fullPath){
        gc $fullPath
    }
}

function Display-ComplianceScanOnly {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Display paths
    get-childitem -recurse \\ed\dst\logs\ComplianceScanOnly-Units\CMVM | ?{$_.name -match $PCName} | select fullname
    Write-Host "`n"
    # Obtain paths
    $fullPath = (get-childitem -recurse \\ed\dst\logs\ComplianceScanOnly-Units\CMVM | ?{$_.name -match $PCName}).FullName
    # For each item, dispay the details
    foreach ($item in $fullPath){
        gc $fullPath
    }
}

function Display-IPUScanOnly {
    param ([Parameter(Mandatory=$true)][string]$PCName)
    
    # Display paths
    get-childitem -recurse \\ed\dst\logs\InPlaceUpgradeStaffScanOnlySDX | ?{$_.name -match $PCName} | select fullname
    Write-Host "`n"
    # Obtain paths
    $fullPath = (get-childitem -recurse \\ed\dst\logs\InPlaceUpgradeStaffScanOnlySDX | ?{$_.name -match $PCName}).FullName
    # For each item, dispay the details
    foreach ($item in $fullPath){
        gc $fullPath
    }
}

function Display-SetupACTLog {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    # Make sure machine is online
    $PCOnline = (Check-PCIsOnline $PCName)
    if ($PCOnline -eq "error"){
        return
    }

    # Full path to file
    $fullPath = Join-Path "\\$PCName\c$" '$WINDOWS.~BT\Sources\Panther\setupact.log'
    # Check the file exists on target machine
    if (Test-Path -Path $fullPath){
        Write-Host "`nDisplaying last 100 lines of $fullPath..`n"
        Get-Content -Tail 100 $fullPath 
    }
    else {
        Write-Host "Unable to display log. Are you sure it's still in the following location? :" -ForegroundColor Red
        Write-Host "$fullPath`n"
    }
    
}

function Copy-PreMigrationScriptMultiple {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    Write-Host "Please make sure that the pre-migration script is in the following location, and that no other items are present in the folder:" -ForegroundColor Yellow
    Write-Host "C:\Workspace\Win10\"

    $question = "`nYou will now be asked to point to the .txt file containing a list of machine names. Continue (y /n) ?"
    do {
        $answer = (Confirm-Answer $question)    
    } until ($answer -eq "y" -or $answer -eq "n")   

    if ($anwer -eq "n") {
        return
    }

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'Text (*.txt)|*.txt'
    }
    $null = $FileBrowser.ShowDialog()

    $hosts = Get-Content $FileBrowser.FileName
    Foreach ($PC in $hosts) {
        $sourcePath = "C:\Workspace\Win10"
        $PCOnline = (Check-ComputerIsOnline $PC)
        if (!($pcOnline -eq $true)) {
            Write-Host "`n$PC is offline" -ForegroundColor Red
            continue        
        }
 
        Else {
            $targetPath = "\\$pc\c$\Workspace\"
            try {
                ROBOCOPY $sourcePath $targetPath  /e /np /w:1 /r:1
                Write-Host "`nSuccessfully copied script to $PC" -ForegroundColor Gray
            }
            catch {
                Write-Host "`nUnable to copy script to $PC!"-ForegroundColor Red
                Write-Host "Cause : $_." -ForegroundColor Red
            }
        }
    }
 }

function Copy-PreMigrationScript {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    $sourcePath = "C:\Workspace\Win10"

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
        $targetPath = "\\$PCName\c$\Workspace\"
        try {
            ROBOCOPY $sourcePath $targetPath  /e /np /w:1 /r:1
            Write-Host "`nSuccessfully copied script to $PCName" -ForegroundColor Gray
        }
        catch {
            Write-Host "`nUnable to copy script to $PCName!"-ForegroundColor Red
            Write-Host "Cause : $_." -ForegroundColor Red
        }        
    }
}

function Check-ComplianceScanRunning{
    param ([Parameter(Mandatory=$true)][string]$PCName)

    gwmi -Cn $PCName win32_process | ?{$_.CommandLine -match "PowerShell"} | Select-Object CommandLine,ProcessId | Format-Custom
}

function Rename-ComputerForWAL {
    
    # Declare OU Path
    $OUPath = "ou=BQ,ou=MVM,ou=ISSD,ou=SDX,ou=UoEX,dc=ed,dc=ac,dc=uk"
    # Show warning
    Write-Host "`nBefore running this command, please make sure the EdLAN DB record is up to date, especially making the following changes :"
    Write-Host "`n1. Change the NetBIOS and DNS name"
    Write-Host "2. Enter the OU path $OUPath"
    Write-Host "`n*** WARNING *** - DO NOT RE-REGISTER THE HOST IN ACTIVE DIRECTORY!" -ForegroundColor Yellow
    Write-Host "`n*** This script will automatically restart the computer you are renaming! ***`n" -ForegroundColor Yellow
    Write-Host "*** Please make sure there is no unsaved data! ***`n" -ForegroundColor Yellow
    # Loop until valid PC name supplied
    do {
        $currentPCName_=read-host "`nPlease enter CURRENT computer name (or enter q to quit) " # Get the old PC name which is to get changed
        # Convert to upper case
        $currentPCName = $currentPCName_.ToUpper()
        If($currentPCName -eq "Q"){ # If entry equals "q" then exit program
            return
        }
        # Check if PC name exists in AD. If so break from loop.
        $validatePCName = (Check-ADComputer $currentPCName)
    } until ($validatePCName -eq "PC exists")

    # Check PC is online
    $PCOnline = (Check-PCIsOnline $currentPCName)
    if ($PCOnline -eq "error"){
        return
    }
    else {
        # Grab current description (if there is one!)
        $currentDescription = Get-ADComputer -Identity $currentPCName -Properties Description | Select-Object -ExpandProperty Description
        if ($currentDescription){
            Write-Host "$currentPCName already has the following description in AD :" -ForegroundColor Yellow
            Write-Host "$currentDescription"
            Write-Host "`n*** WARNING *** This will be overwritten!`n" -ForegroundColor Yellow
        }

        # Check to make sure no one is logged in
        # Declare array to store logged in usernames
        $loggedInusers = @()

        # Get windows explorer processes from remote computer
        $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ComputerName $currentPCName -ErrorAction SilentlyContinue)
        # If no explorer processes are running then most likely no one is logged in
        if ($explorerprocesses.Count -eq 0) {
            "Currently, no one appears to be logged into $currentPCName."
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
            Write-Host "`nThe following users have explorer.exe process running on $currentPCName, which indicates that they are currently logged in :"
            Write-Host "`n$displayArray`n"
            Write-Host "*** WARNING *** - Renaming the computer will force a restart, causing any unsaved work to be lost." -ForegroundColor Yellow
            # Loop until valid reply
            do {
                $question = "`nDo you wish to continue (y/n) ?"
                $answer = (Confirm-Answer $question)
            } until ($answer -ne "error")

            if ($answer -eq "n"){
                return
            }
        }

        # Get details of what the computer is to be renamed to
        $newPCName_ = Read-Host "`nPlease enter NEW PC name (or enter q to quit) "
        $newPCName = $newPCName_.ToLower()                
        If($newPCName -eq "q"){ # If entry equals "q" then exit program
            return
        }
        # Make sure name is 13 characters
        if($newPCName.Length -ne 13){
            # Name is not the typical length of 13 characters (MVM-BIOQ-1234)
            Write-Host "`nThe new name you have entered ($newPCName) does not appear to follow the typical Windows 10 naming convention, which should be 13 characters." -ForegroundColor Yellow
            do {
                $newPCquestion = "`nDo you wish to proceed anyway (y/n) ? "
                $newPCAnswer = (Confirm-Answer $newPCquestion)     
            } until ($newPCAnswer -ne "error")
            # If user doesn't want to proceed then quit the script
            if ($newPCAnswer -eq "n"){
                return
            }
        }
        # Make sure the new name doesn't exist
        $PCExists = (Check-ADComputer $newPCName)
        if ($PCExists -eq "error") {        
            Write-Host "`nComputer name $newPCName is available`n" -ForegroundColor Green
        }
        else {        
            Write-Host ""`n"There is already a computer in Active Directory with this name. Please enter a different name.`n" -ForegroundColor Red
            return
        }
        # Get confirmation and perform the change
        do {
            $confirmQuestion = "`nAre you sure you wish to change $currentPCName to $newPCName (y / n) ? "
            $confirm = (Confirm-Answer $confirmQuestion)
        } until ($confirm -ne "error")
            
        if ($confirm -eq "y") {
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
                # Set description
                Write-Host "`nAttempting to add description to computer AD object..."
                # Prepare string to add to description
                $description = $currentPCName + ':SDX:STAFF:ed.ac.uk/UoEX/SDX/ISSD/MVM/BQ:'
                # For testing purposes
                #$description = $currentPCName + ':SDX:STAFF:ed.ac.uk/UoEX/SDX/ISSD/IS_DaaS/IS/USD:'
                # Attempt to add description. Waiting for 10 seconds to make sure AD object has been renamed.
                $setDescr = $False
                $adwait = 0
                Write-Host "Checking for AD object $newPCName"
                WHILE ($adwait -lt 10) {
                    Write-Host $adwait
                    Try { 
                        If (Get-AdComputer $newPCName -ErrorAction Stop){
                            $setDescr=$True
                            $adwait=10
                        }
                    }
	                Catch { 
                        $adwait++
                        Start-Sleep -s 5
                    }
                }
                If ($setDescr -eq $True) {
                    # Make sure name has been changed in AD
                    try {
                        Set-ADComputer $newPCName -Description $description -ErrorAction Stop
                        Write-Host "`n*** Description appears to have been set successfully. Probably best to double check the AD object to make sure. ***`n" -ForegroundColor Yellow
                    }
                    catch {
                        Write-Host $_.
                        Write-Host "`n*** ERROR *** - Unable to set description on $newPCName. Please do so manually.`n" -ForegroundColor Red
                    }
                }                                    
                                    
        }
        elseIf ($confirm -eq "n") {        
            Return
        }          
    }
}

function Check-WindowsUpdates{
    param ([Parameter(Mandatory=$true)][string]$PCName)

    get-wmiobject -ComputerName $PCName -query "SELECT * FROM CCM_SoftwareUpdate" -namespace "ROOT\ccm\ClientSDK" | Select Name, ArticleID, PercentComplete
}

function Install-WindowsUpdates{
    param ([Parameter(Mandatory=$true)][string]$PCName)
    # requires a list of Article IDs to be entered. Article IDs can be obtined by running "Check-WindowsUpdates"

    get-wmiobject -ComputerName $PCName -query "SELECT * FROM CCM_SoftwareUpdate" -namespace "ROOT\ccm\ClientSDK" | ? ArticleID -in @('2461484', '3000483')| % { ([wmiclass]'ROOT\ccm\ClientSDK:CCM_SoftwareUpdatesManager').InstallUpdates($_)};
}

function Check-ForPendingCCMReboot {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    ([wmiclass]"\\$PCName\ROOT\ccm\ClientSDK:CCM_ClientUtilities").DetermineIfRebootPending().RebootPending
}

function Remove-OLERegistryEntry {
    

    <#
        .Synopsis
        Removes the following registry key : HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Ole

        .Description
        Backs up the following registry key to C:\Workspace\ and then removes the contents from the original location :
        
        HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Ole

        Removing this key hopefully resolves the issue of EliteDesk 800 G1 models from failing the Windows 10 upgrade and rolling back to Windows 7 during the process.
        
        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

    #>
    
    Write-Host "`nThis tool attempts to backup the following registry key to C:\Workspace on a remote machine`n"
    Write-Host "`tHKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Ole`n" -ForegroundColor Yellow
    Write-Host "and then delete the above keys contents from the original location.`n"

    # Ask for PC name
    $PCName = Read-Host "Enter computer name on which remove the key "

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
    Write-Host "`n$PCName appears to be online...`n" -ForegroundColor Green

    # Check it's a G1
    $model = (Get-WmiObject Win32_ComputerSystem -ComputerName $PCName).Model

    if ($model -like "*G1*" -or $model -like "*7010" -or $model -like "Lattitude*" -or $model -like "*8300"){
        Write-Host "$PCName appears to be a model where the rollback issue may occur ($model)"
    }
    else{
        Write-Host "$PCName doesn't appear to be a model where the rollback issue may occur ($model). Exiting script...`n" -ForegroundColor Red
        return
    }   
  
    # Create new PS Session
    $Session = New-PSSession $PCName

    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName)

        # Nested function
        function RegKey-Exists{
            param ([Parameter(Mandatory=$true)][string]$regKey)

            # Find key
            Write-Host "Attempting to find key...`n"
            if (Get-Item -Path $key -ErrorAction SilentlyContinue){
                Write-Host "$key exists.`n" -ForegroundColor Green
                return "exists"
            }
            else {
                Write-Host "$key does not exist.`n" -ForegroundColor Red
                return "error"
            }

        }
                
        # Declare key location
        $key = 'HKLM:\SOFTWARE\Microsoft\Ole\'                
        
        # Declare Extensions key location
        $extKey = 'HKLM:\SOFTWARE\Microsoft\Ole\Extensions\'

        # Check if key exists. If not then quit.
        $keyExists = (RegKey-Exists $key)
        Write-Host $keyExists
        if ($keyExists -eq "error"){
            Write-Host "Exiting script..."
            return   
        }

        # Backup key
        Write-Host "Backing up $key to C:\Workspace...`n"
        $temp_key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Ole" 
        REG EXPORT $temp_key "C:\Workspace\Ole.reg"

        # Check for Extensions subkey
        $extKeyExists = (RegKey-Exists $extKey)
        Write-Host $extKeyExists
        if ($keyExists -eq "exists"){
            Write-Host "This script has detected that the Extensions subkey at the following location exists:" -ForegroundColor Yellow
            Write-Host "`n`t$extKey" -ForegroundColor DarkYellow
            Write-Host "`nTypically this cannot be removed using this script.`nYou will need to manually take ownership of this subkey to delete.`n" -ForegroundColor Yellow
        }

        # Remove values and sub keys
        Write-Host "Removing values...`n"
        Remove-ItemProperty -Path $key -Name * -Exclude "Default"
        Write-Host "Removing subkeys...`n"
        Remove-Item -Path $key* -Recurse

    }  -ArgumentList $PCName

    # Remove PS Session
    Remove-PSSession $Session

    Write-Host "Done!`n" -ForegroundColor Green

}

function Display-IsUoESDMember {
    param ([Parameter(Mandatory=$true)][string]$PCName)
    
    (get-adcomputer $PCName -property canonicalname).canonicalName -match "^ed.ac.uk/UoESD/"

}

function Display-InESUGroup{
    param ([Parameter(Mandatory=$true)][string]$PCName)

    (Get-Adcomputer $PCName -Property memberOf).MemberOf -contains "CN=ESU MVM Machines,OU=ESU,OU=DST,OU=Auth,OU=UoEX,DC=ed,DC=ac,DC=uk"
}

function Display-ESULog {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    Get-ChildItem -recurse \\ed\dst\logs\ESU |?{$_.Basename -eq "$PCName"}| % {Write-Host $_.Fullname ;Get-Content $_.Fullname}
}

function Display-ESOGPStatus{
    param ([Parameter(Mandatory=$true)][string]$PCName)

(gpresult /r /s $PCName /SCOPE COMPUTER /Z).Split("`n") |Select-String -context 0,7 "GPO: ESU - Configuration"
}

function Display-ESULogs{
    $(foreach ($log in (Get-ChildItem -recurse \\ed\dst\logs\esu -File |?{$_.Name -match "(MVM|CTR|BQ-PC)"}|Group-Object Basename)){If ($log.count -gt 1) {$log.Group |Sort Lastwritetime|Select -last 1}  else {$log.Group}})  |Select BaseName,DirectoryName,@{Name="Date";Expression={Get-Date $_.LastWriteTime -Format "dd/MM/yyyy"}},FullName|out-gridview -Title "ESU Status"
}