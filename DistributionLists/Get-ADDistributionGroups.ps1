<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER SourceDataPath

.EXAMPLE

.NOTES
    Version:        1.0
    Author:         Brian Jones
    Creation Date:  2025-01-17
    Last Modified:  2025-01-17

#>

param (
    [parameter(Mandatory = $true)]
    [string]$ExportPath
)

Import-Module ActiveDirectory
$ExportPath = $ExportPath
$DistributionGroups = Get-ADGroup -Filter "GroupCategory -eq 'Distribution'" -Properties *
foreach($Group in $DistributionGroups){
    try{
        $GroupMembers = Get-ADGroupMember -Identity $Group.DisplayName
    } catch {
        Write-Host "No members found in $Group." -ForegroundColor Yellow
        Continue
    }

    foreach($Member in $GroupMembers){
        $OutputObj = [PSCustomObj]@{
            GroupName = $Group.DisplayName
            GroupEmail = $Group.Mail
            Name = $Member.Name
        }
    }
}

$DistributionGroups | Export-CSV $ExportPath