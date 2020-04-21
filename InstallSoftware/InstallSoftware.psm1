function Get-MsiInformation {
    param ([Parameter(Mandatory=$true)][string]$Path)

    [string[]]$Property = ( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )
    $MsiFile = Get-Item -Path $Path
    Write-Host "Executing on $P"
                   
    # Read property from MSI database
    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                   
    # Build hashtable for retruned objects properties
    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName}
    ForEach ( $Prop in $Property ) {
        Write-Verbose -Message "Enumerating Property: $Prop"
        $Query = "SELECT Value FROM Property WHERE Property = '$( $Prop )'"
        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
 
        # Return the value to the Property Hash
        $PSObjectPropHash.Add($Prop, $Value)
    }
          
    # Build the Object to Return
    $Object = @( New-Object -TypeName PSObject -Property $PSObjectPropHash )
                   
    # Commit database and close view
    $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
    $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
    $MSIDatabase = $null
    $View = $null          
    Write-Output -InputObject @( $Object )
                   
    # Run garbage collection and release ComObject
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
    [System.GC]::Collect()    
}


function Install-RStudio {

    <#
        .Synopsis
        Installs R & RStudio

        .Description
        Installs R and RStudio on a remote computer. 
        
        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\RStudio\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following versions are installed :

        R-3.6.1
        RStudio-1.2.1335

    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\RStudio\"
    # Get R filename
    $R_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "R-*").Name
    # Get RStudio filename
    $RStudio_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "RStudio*").Name

    # Show disclaimer
    Write-Host "`nPlease note that this script will install versions of R and RStudio from the R website, not the versions in Software Center. It will install the following versions :" -ForegroundColor Yellow
    Write-Host "`n$R_`n$RStudio_`n"
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install R & RStudio "

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
    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "RStudio-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying $R_ and $RStudio_ to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $R = $targetPath.ToString() + "\" + $R_.ToString()
    $RStudio = $targetPath.ToString() + "\" + $RStudio_.ToString()
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $R, $RStudio,$R_, $RStudio_)

        # Install R
        Write-Host "`nInstalling $R_ on $PCName...."  
        Start-Process -wait -FilePath $R -ArgumentList "/VERYSILENT"
        # Install RStudio
        Write-Host "Installing $RStudio_ on $PCName...."
        Start-Process -wait -FilePath $RStudio -ArgumentList "/S"
        Write-Host "`nR and RStudio installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $R, $RStudio, $R_, $RStudio_

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    Write-Host "`nDone!`n" -ForegroundColor Green
}



function Install-GraphPadPrism8 {

    <#
        .Synopsis
        Installs Graphpad Prism 8

        .Description
        Installs GraphPad Prism 8 on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Prism\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed : Prism 8.2.0.435

        *** End user on remote computer will need to input license details on 1st run ***

    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Prism\"
    # Get Prism filename
    $Prism_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "InstallPrism*").Name
    
    # Show disclaimer
    Write-Host "`nPlease note that this script will install Prism 8.4.0.671, not the version in Software Center.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to Prism "

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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "Prism-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying Prism to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $Prism = $targetPath.ToString() + "\" + $Prism_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $Prism)

        # Install Prism
        Write-Host "`nInstalling $Prism on $PCName...."  
        Start-Process -wait -FilePath $Prism -ArgumentList "/QUIET"
        
        Write-Host "`nPrism installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $Prism

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    # Set permissions on shortcut so it can be deleted
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\GraphPad Prism 8.lnk"
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Modify","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-SnapGeneViewer4 {

    <#
        .Synopsis
        Installs SnapGene Viewer 4

        .Description
        Installs SnapGene Viewer 4 on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\SnapGeneViewer\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  4.3.10

    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\SnapGeneViewer\"
    # Get R filename
    $SGV_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "snapgene_viewer*").Name
    
    # Show disclaimer
    Write-Host "`nPlease note that this script will install SnapGeneViewer 4.3.10.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to SnapGeneViewer "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "SGViewer-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying SnapGeneViewer to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $SGV = $targetPath.ToString() + "\" + $SGV_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $SGV)

        # Install Prism
        Write-Host "`nInstalling $SGV on $PCName...."  
        Start-Process -wait -FilePath $SGV -ArgumentList "/S"
        
        Write-Host "`nSnapGene Viewer installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $SGV

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    # Create shortcut and place on Desktop of all users
    Write-Host "`nCreating Desktop shortcut to C:\Program Files (x86)\SnapGene Viewer\SnapGene Viewer.exe"
    $targetFile = "\\$PCName\c$\Program Files (x86)\SnapGene Viewer\SnapGene Viewer.exe"
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\SnapGene Viewer.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WscriptShell.CreateShortcut($shortcutFile)
    $Shortcut.TargetPath = $targetFile
    $Shortcut.Save()

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-FCSExpress7 {

    <#
        .Synopsis
        Installs FCS Express 7

        .Description
        Installs FCS Express 7 Research Installation on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\FCSExpress\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  4.3.10

        *** This is a demo version only! For full version user will need to input license details. ***
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\FCSExpress\"
    # Get FCSExpress filename
    $FCSExpress_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "FCSExpress*").Name
    
    # Show disclaimer
    Write-Host "`nPlease note that this script will install FCSExpress 7 Research Installation 7.04.0014.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to FCSExpress "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "FCSExpress-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying FCSExpress to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $FCSExpress = $targetPath.ToString() + "\" + $FCSExpress_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $FCSExpress)

        # Install Prism
        Write-Host "`nInstalling $FCSExpress on $PCName...."  
        Start-Process -wait -FilePath $FCSExpress -ArgumentList "/SP- /VERYSILENT"
        
        Write-Host "`nFCSExpress 7 Research Install installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $FCSExpress

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    # Create shortcut and place on Desktop of all users
    Write-Host "`nCreating Desktop shortcut to C:\Program Files\De Novo Software\FCS Express 7 Research Edition"
    $targetFile = "\\$PCName\c$\Program Files\De Novo Software\FCS Express 7 Research Edition\FCS Express App.exe"
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\FCS Express 7.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WscriptShell.CreateShortcut($shortcutFile)
    $Shortcut.TargetPath = $targetFile
    $Shortcut.Save()

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-Docker {
    <#
        .Synopsis
        Installs Docker application on a remote computer.

        .Description
        Installs Docker application on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Docker\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  2.0.10625
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Docker\"
    # Get Docker filename
    $Docker_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "InstallDock*").Name
  
    # Show disclaimer
    Write-Host "`nPlease note that this script will install Docker version 2.0.10625.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install Docker "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "Docker-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying Docker to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $Docker = $targetPath.ToString() + "\" + $Docker_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $Docker)

        # Install Prism
        Write-Host "`nInstalling $Docker on $PCName...."  
        Start-Process -wait -FilePath $Docker -ArgumentList "/quiet"
        
        Write-Host "`nDocker installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $Docker

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    # Get shortcut location
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\Docker for Windows.lnk"

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-ImageJ {
    <#
        .Synopsis
        Installs ImageJ application on a remote computer.

        .Description
        Installs ImageJ application on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Fiji\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Fiji\"
    
  
    # Show disclaimer
    Write-Host "`nPlease note that this is a stand alone application, not an install exe." -ForegroundColor Yellow
    Write-Host "It will copy the application to the following location on the remote machine :`n"
    Write-Host "`tC:\Users\Public\Documents\`n"
    Write-Host "and create a shortcut on the desktop for all users."
    Write-Host "`nPlease note this may also take some time!`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install ImageJ "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath = "\\$PCName\c$\Users\Public\Documents\"
    
    # Copy the install packages to remote machine
    Write-Host "Copying ImageJ / Fiji to $targetPath"
    ROBOCOPY $softwareLocation $targetPath Fiji.app.zip /njh /njs /ndl /nc /E /R:1 /W:5
    
    # Unzip package
    $sourcePath = "\\$PCName\c$\Users\Public\Documents\Fiji.app.zip"
    $targetExtract = "\\$PCName\c$\Users\Public\Documents\"
    #Extract .zip
    Write-Host "`nExtracting Fiji.app.zip......"
    expand-archive -path $sourcePath -DestinationPath $targetExtract -Force

    # Create shortcut and place on Desktop of all users
    Write-Host "`nCreating Desktop shortcut to C:\Users\Public\Documents\Fiji.app\ImageJ-win64.exe...."
    $targetFile = "\\$PCName\c$\Users\Public\Documents\Fiji.app\ImageJ-win64.exe"
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\ImageJ.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WscriptShell.CreateShortcut($shortcutFile)
    $Shortcut.TargetPath = $targetFile
    $Shortcut.Save()

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    # Remove .zip file
    Write-Host "`nDeleting Fiji.app.zip..."
    Remove-Item –path $sourcePath –recurse

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-ImageStudioLite {
    <#
        .Synopsis
        Installs Docker application on a remote computer.

        .Description
        Installs Docker application on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\ImageStudioLite\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  2.0.10625
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\ImageStudioLite\"
    # Get Software filename
    $ImageStudioLite_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "Win_Image*").Name
  
    # Show disclaimer
    Write-Host "`nPlease note that this script will install Image Studio Lite version 5.2.5.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install Image Studio Lite "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "ISL-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying Image Studio Lite to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $ImageStudioLite = $targetPath.ToString() + "\" + $ImageStudioLite_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $ImageStudioLite)

        # Install Prism
        Write-Host "`nInstalling $ImageStudioLite on $PCName...."  
        Start-Process -wait -FilePath $ImageStudioLite -ArgumentList "/VERYSILENT"
        
        Write-Host "`nImageStudioLite installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $ImageStudioLite

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    # Get shortcut location
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\Image Studio Lite Ver 5.2.lnk"

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-FlowJo {

    <#
        .Synopsis
        Installs FlowJo

        .Description
        Installs FlowJo 10.6.1 on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\FlowJo\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed : FlowJo 10.6.1

        *** End user on remote computer will need to input license details on 1st run ***

    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\FlowJo\"
    # Get FloJo filename
    $FlowJo_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "FlowJo-*").Name
    
    # Show disclaimer
    Write-Host "`nPlease note that this script will install FlowJo 10.6.1.`n" -ForegroundColor Yellow
    Write-Host "License details will need to be entered on first run!`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install FlowJo "

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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "FlowJo-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying FlowJo to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $FlowJo = $targetPath.ToString() + "\" + $FlowJo_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $FlowJo)

        # Install FlowJo
        Write-Host "`nInstalling $FlowJo on $PCName...."  
        Start-Process -wait -FilePath $FlowJo -ArgumentList "-i silent"
        
        Write-Host "`nFlowJo installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $FlowJo

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    Write-Host "`nDone!`n" -ForegroundColor Green
    Write-Host "Please note that no shortcut is created on the Desktop, programs will need to be started from the start menu.`n" -ForegroundColor Yellow
}

function Install-DSSPlayerLite {
    <#
        .Synopsis
        Installs DSS Player Lite

        .Description
        Installs DSS Player Lite on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Olympus DSS Player\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed : DSS Player Lite 2.1.1
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Olympus DSS Player\DSSPlayerLite"    
    
    # Show disclaimer
    Write-Host "`nPlease note that this script will install DSS Player Lite 2.1.1`n" -ForegroundColor Yellow
     
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install DSS Player Lite "

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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "DSSPlayerLite-$timestamp" -ItemType "directory"
    
    $targetPath = $targetPath.ToString()
    # Copy the install packages to remote machine
    Write-Host "Copying DSS Player Lite to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /E /R:1 /W:5

    # Set path to .zip
    $sourcePath = "$targetPath" + "\" + "DSS_PLayer_Lite-2.1.1.zip"
    # Just use $targetPath location as extraction location
    Write-Host "`nExtracting DSS_Player_Lite-2.1.1.zip......"
    expand-archive -path $sourcePath -DestinationPath $targetPath -Force

    # Get full path to setup.exe
    $DSS = "$targetPath" + "\" + "DSS_Player_Lite-2.1.1" + "\" + "Setup.exe"

    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $DSS)

        # Install DSS Player Lite
        Write-Host "`nInstalling $DSS on $PCName...."  
        Start-Process -wait -FilePath $DSS -ArgumentList "/s"
        
        Write-Host "`nDSS Player Lite installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $DSS


    # Create shortcut and place on Desktop of all users
    Write-Host "`nC:\Program Files (x86)\Olympus\DSSPlayerLite\DSSPly.exe"
    $targetFile = "C:\Program Files (x86)\Olympus\DSSPlayerLite\DSSPly.exe"
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\DSS Player Lite.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WscriptShell.CreateShortcut($shortcutFile)
    $Shortcut.TargetPath = $targetFile
    $Shortcut.Save()

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    # Remove .zip file
    Write-Host "`nDeleting DSS_PLayer_Lite-2.1.1.zip..."
    Remove-Item –path $sourcePath –recurse

    Write-Host "`nDone!`n" -ForegroundColor Green
    
}

function Install-iSpy {
    <#
        .Synopsis
        Installs iSpy application on a remote computer.

        .Description
        Installs iSpy application on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\ImageStudioLite\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  2.0.10625
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\iSpy\"
    # Get Software filename
    $iSpy_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "iSpy*").Name
  
    # Show disclaimer
    Write-Host "`nPlease note that this script will install iSpy version 7.2.1.0.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install iSpy "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "iSpy-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying iSpy installer to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $iSpy = $targetPath.ToString() + "\" + $iSpy_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $iSpy)

        # Install iSpy
        Write-Host "`nInstalling $iSpy on $PCName...."  
        Start-Process -wait -FilePath $iSpy -ArgumentList "/QUIET"
        
        Write-Host "`niSpy installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $iSpy

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    # Get shortcut location
    $shortcutFile = "\\$PCName\c$\Users\Public\Desktop\iSpy (64 bit).lnk"

    # Set permissions on shortcut so that anyone can delete
    Write-Host "`nSetting permissions on desktop shortcut......"
    $acl = Get-Acl $shortcutFile
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone","Delete","Allow")
    $acl.SetAccessRule($accessRule)
    $acl | Set-Acl $shortcutFile

    Write-Host "`nDone!`n" -ForegroundColor Green
}


function Install-GeneiousPrime2019 {
    <#
        .Synopsis
        Installs GeneiousPrime2019 application on a remote computer.

        .Description
        Installs GeneiousPrime2019 application on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\GeneiousPrime2019\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  2.3
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\GeneiousPrime\2019\"
    # Get Software filename
    $GP_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "Geneious*").Name
  
    # Show disclaimer
    Write-Host "`nPlease note that this script will install Geneious Prime 2019 version 2.3.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install Geneious Prime "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "GeneiousPrime-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying Geneious Prime installer to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $GP = $targetPath.ToString() + "\" + $GP_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $GP)

        # Install iSpy
        Write-Host "`nInstalling $GP on $PCName...."  
        Start-Process -wait -FilePath $GP -ArgumentList "/QUIET"
        
        Write-Host "`nGeneious Prime installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $GP

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-VirtualBox {
    <#
        .Synopsis
        Installs VirtualBox application on a remote computer.

        .Description
        Installs VirtualBox application on a remote computer.

        REQUIREMENTS:
        - AD object for remote PC exists
        - Remote PC is online
        - Read access permissions at \\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\VirtualBox\
        - Must be executed from computer on trusted subnet
        - Administrator permissions on remote computer

        The following version is installed :  2.3
    #>

    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\VirtualBox"
    # Get Software filename
    $VB_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "VirtualBox*").Name
  
    # Show disclaimer
    Write-Host "`nPlease note that this script will install VirtualBox version 6.1.2.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install VirtualBox "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "VirtualBox-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying VirtualBox installer to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $VB = $targetPath.ToString() + "\" + $VB_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $VB)

        # Install VirtualBox
        Write-Host "`nInstalling $VB on $PCName...."  
        Start-Process -wait -FilePath $VB -ArgumentList "--silent"
        
        Write-Host "`nVirtualBox installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $VB

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    Write-Host "`nDone!`n" -ForegroundColor Green
}

function Install-VNCViewer {
    # Declare locaiton of install packages
    $softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\VNCViewer"
    # Get Software filename
    $VNCViewer_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "VNC*").Name
  
    # Show disclaimer
    Write-Host "`nPlease note that this script will install VNCViewer version 6.20.113.`n" -ForegroundColor Yellow
    
    # Loop until valid answer is given
    do {
        $confirm = "Do you wish to continue (y/n) ?"
        $confirmAnswer = (Confirm-Answer $confirm)
    } until ($confirmAnswer -ne "error")
    # If answer is n then quit
    if ($confirmAnswer -eq "n"){
        return
    }
    
    # Ask for PC name
    $PCName = Read-Host "`nEnter computer name on which to install VNCViewer "

    # Make sure PC exists in AD
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

    # Declare target path for install packages
    $targetPath_ = "\\$PCName\c$\Windows\Temp\"
    # Get timestamp
    $timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
    # Create temporary folder with timestamp as name so we are sure it's a unique name
    $targetPath = New-Item -Path $targetPath_ -Name "VNCViewer-$timestamp" -ItemType "directory"
    # Copy the install packages to remote machine
    Write-Host "Copying VNCViewer installer to $targetPath ...."
    ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
    # Get full paths to install packages on remote machine
    $VNCViewer = $targetPath.ToString() + "\" + $VNCViewer_.ToString()
    
    # Create new PS Session
    $Session = New-PSSession $PCName
    # Invoke command on the remote machine
    Invoke-Command -Session $Session -ScriptBlock {
        param($PCName, $VNCViewer)

        # Install VNCViewer
        Write-Host "`nInstalling $VNCViewer on $PCName...."  
        Start-Process -wait -FilePath $VNCViewer -ArgumentList "/quiet"
        
        Write-Host "`nVNCViewer installed." -ForegroundColor Green

    }  -ArgumentList $PCName, $VNCViewer

    # Remove install files
    Write-Host "`nRemoving install files..."
    Remove-Item –path $targetPath –recurse
    # Remove PS Session
    Remove-PSSession $Session

    Write-Host "`nDone!`n" -ForegroundColor Green
}