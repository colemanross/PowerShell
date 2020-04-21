
# Known issues with *-LocalGroupMember commands
# https://github.com/PowerShell/PowerShell/issues/2996
# https://superuser.com/questions/1131901/get-localgroupmember-generates-error-for-administrators-group
function Display-LocalGroupMembers{
    param(
        [Parameter(Mandatory=$true)][string]$PCName,
        [Parameter(Mandatory=$true)][string]$groupName
    )
    $question = "Yes or no?"

    $answer = (Confirm-Answer $question)

    Write-Host "The answer is $answer"
}