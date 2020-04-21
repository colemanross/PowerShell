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

function Check-PCisOnline {
    param ([Parameter(Mandatory=$true)][string]$PCName)

    try {
        if ([bool] (Test-Connection $PCName -Quiet -count 1 -ErrorAction Stop)) {
            return "PC online"
        }
    } 
    catch {
        Write-Host "`n$PCName is offline!`n" -ForegroundColor Red
        return "error"
    }
}

# Declare locaiton of install packages
$softwareLocation = "\\cmvm.datastore.ed.ac.uk\cmvm\cmvm\shared\IT\Software\PC\Prism\"
# Get Prism filename
$Prism_ = (Get-ChildItem $softwareLocation | Where-Object -Property Name -like "InstallPrism*").Name
    
# Show disclaimer
Write-Host "`nPlease note that this script will install Prism 8.2.0.435, not the version in Software Center.`n" -ForegroundColor Yellow

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
# Check PC is online
$PCOnline = (Check-PCIsOnline $PCName)
if ($PCOnline -eq "error"){
    return
}
Write-Host "`n$PCName appears to be online...`n" -ForegroundColor Green

# Declare target path for install packages
$targetPath_ = "\\$PCName\c$\Workspace\"
# Get timestamp
$timestamp = Get-Date -f MM-dd-yyy_HH_mm_ss
# Create temporary folder with timestamp as name so we are sure it's a unique name
$targetPath = New-Item -Path $targetPath_ -Name "Prism-$timestamp" -ItemType "directory"
# Copy the install packages to remote machine
Write-Host "Copying Prism to $targetPath ...."
ROBOCOPY $softwareLocation $targetPath /njh /njs /ndl /nc /ns /E /R:1 /W:5
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
Write-Host "`nDone!`n" -ForegroundColor Green