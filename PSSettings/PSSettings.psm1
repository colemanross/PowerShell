function Display-Commands{
    #Get-Module | Where-Object {($_.ModuleType -like "Script") -and ($_.Name -notlike "NetAdapter.Format.Helper") -and ($_.Name -notlike "PSReadline")} | select Name, ExportedFunctions | fl
    Get-ChildItem function:\ | Where-Object {($_.Source -notlike "") -and ($_.Source -notlike "Microsoft.PowerShell.Utility") -and ($_.Source -notlike "PSReadline") -and ($_.Source -notlike "*.hlo") -and ($_.Source -notlike "DnsClient")} | Select Name, Source 
}

function Display-FunctionCode{
    param ([Parameter(Mandatory=$true)][string]$functionName)
    $ErrorActionPreference = "Stop"

    try {
        (Get-Command $functionName).Definition
    }
    catch{
        Write-Host "`nCannot find function name $functionName!. Are you certain this is the name of a function?" -ForegroundColor Red
        Write-Host "For a list of function names, enter the following command:" -ForegroundColor Yellow
        Write-Host "`n`tDisplay-Commands`n"
    }

}

# Sends a CTRL+C to the process
function Quit-Process {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [System.Windows.Forms.SendKeys]::SendWait("^{c}")
}

function Reload-Modules {
    Write-Host "Re-importing PS Custom Modules..."
    Write-Host "ADActions..."
    Import-Module ADActions -force -WarningAction SilentlyContinue
    Write-Host "ADInfo..."
    Import-Module ADInfo -force -WarningAction SilentlyContinue
    Write-Host "ADPermissions..."v
    Import-Module ADPermissions -force -WarningAction SilentlyContinue
    Write-Host "CreateADGroups..."
    Import-Module CreateADGroups -force -WarningAction SilentlyContinue
    Write-Host "PCActions..."
    Import-Module PCActions -force -WarningAction SilentlyContinue
    Write-Host "PCInfo..."
    Import-Module PCInfo -force -WarningAction SilentlyContinue
    Write-Host "PCPermissions..."
    Import-Module PCPermissions -force -WarningAction SilentlyContinue
    Write-Host "PSSettings..."
    Import-Module PSSettings -force -WarningAction SilentlyContinue
    Write-Host "ValidateInput..."
    Import-Module ValidateInput -force -WarningAction SilentlyContinue
    Write-Host "Win7Migration..."
    Import-Module Win7Migration -force -WarningAction SilentlyContinue
    Write-Host "Office365..."
    Import-Module Office365 -force -WarningAction SilentlyContinue
    Write-Host "Done!"
}

# timestamp the log file
function Timestamp {
    param ([Parameter(Mandatory=$true)][string]$logfile)

    $timestamp = Get-Date -format F
    Add-Content -value "--------------- $timestamp ---------------" -path $logfile
}