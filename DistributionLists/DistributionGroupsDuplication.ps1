<#
.SYNOPSIS
    This script takes an input from a CSV which contains and export of distribution lists from AD and creates those Distribution Lists in M365.

.DESCRIPTION
    This script take the exported data from AD and does the following:
    - Checks to see if the group already exists in M365 EXO, if not it will create the group.
    - Tries to map users from AD against users in EXO.
    - Adds matched users to New and Existing Distribution Groups.

    Before using this script you must connect to EXO

    Connect-ExchangeOnline `
        -CertificateThumbPrint $CertThumbprint `
        -AppID $ClientId `
        -Organization $Organization

.PARAMETER SourceDataPath
    The path to the CSV Export from AD.

.EXAMPLE
    DistributionGroupsDuplication -SourceDataPath c:\temp\myfile.csv

.NOTES
    Version:        1.0
    Author:         Brian Jones
    Creation Date:  2025-01-17
    Last Modified:  2025-01-17

#>

param (
    [parameter(Mandatory = $true)]
    [string]$SourceDataPath
)

function Get-Destination-DistributionGroups {
    [CmdletBinding()]
    param ()

    $groupMembers = @()
    Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
        $group = $_
        try {
            $members = Get-DistributionGroupMember -Identity $group.Identity -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to retrieve members for group '$($group.DisplayName)': $_"
            return
        }
        
        if ($members.Count -eq 0) {
            $groupMembers += [PSCustomObject]@{
                GroupName   = $group.DisplayName
                GroupEmail  = $group.PrimarySmtpAddress
                MemberName  = "No members"
                MemberEmail = "No members"
            }
        }
        else {
            $members | ForEach-Object {
                $member = $_
                $groupMembers += [PSCustomObject]@{
                    GroupName   = $group.DisplayName
                    GroupEmail  = $group.PrimarySmtpAddress
                    MemberName  = $member.DisplayName
                    MemberEmail = $member.PrimarySmtpAddress
                }
            }
        }
    }
    return $groupMembers
}

function Compare-DistributionGroups(){
    param (
        [Parameter(Mandatory=$true)]
        [object]$DestinationDistributionGroups,
        [Parameter(Mandatory=$true)]
        [object]$SourceDistributionGroups
    )

    $matchedGroup = $null
    $matchedGroups = @()
    $unMatchedGroups = @()
    foreach($item in $SourceDistributionGroups){
        $matchedGroup = $DestinationDistributionGroups | Where GroupName -eq $item.GroupName
        if($matchedGroup){
            $matchedGroups += $matchedGroup
        } else {
            $unMatchedGroups += $item
        }
    }

    $results = [PSCustomObject]@{
        Matched = $matchedGroups
        UnMatched = $unMatchedGroups
    }
    return $results
}

function Add-Remove-Users-To-Groups(){
    param (
        [Parameter(Mandatory=$true)]
        [object]$SourceDistributionGroups
    )

    $DestinationGroup = $null
    $DestinationGroupMembers = $null

    $uniqueGroupName = $SourceDistributionGroups | sort GroupName -Unique
    foreach($group in $uniqueGroupName){
        $DestinationGroup = Get-DistributionGroup $group.GroupName
        if($DestinationGroup){
            $SourceGroupMembers = $SourceDistributionGroups | Where GroupName -EQ $group.GroupName
            $DestinationGroupMembers = Get-DistributionGroupMember -Identity $group.GroupName
            if(!$DestinationGroupMembers){
                foreach($member in $SourceGroupMembers){
                    try {
                        $user = Get-User $member.Name -ErrorAction Stop
                    } Catch {
                        $global:usersNotFound += $member.Name
                        continue
                    }
                    
                    try {
                        Add-DistributionGroupMember -Identity $DestinationGroup.DisplayName -Member $member.Name -ErrorAction Stop
                        $global:addedToDistributionGroup += $member.Name, $DestinationGroup.Name
                    } catch {
                        Write-Host "Failed to add user."
                        Break
                    }
                }
            } else {
                $comparison = Compare-Object -ReferenceObject $DestinationGroupMembers -DifferenceObject $SourceGroupMembers -Property Name
                foreach($lineItem in $comparison | where SideIndicator -EQ "=>"){
                    try {
                        $user = Get-User $lineItem.Name -ErrorAction Stop
                    } Catch {
                        $global:usersNotFound += $lineItem.Name
                        continue
                    }

                    try {
                        Add-DistributionGroupMember -Identity $DestinationGroup.Name -Member $user -ErrorAction Stop
                        $loggingString = $line.Name + ";" + $DestinationGroup.Name
                        $global:addedToDistributionGroup += $line.name, $DestinationGroup.Name
                    } catch {
                        Write-Host "Failed to add user."
                        Break
                    }
                }
            }
            
        } else {
            Write-Output $DestinationGroup.DisplayName "does not exists."
        }
    }
}

$global:usersNotFound = @()
$global:addedToDistributionGroup = @()
$global:createdDistributionGroup = @()
$global:failedToCreateDistributionGroup = @()

#Import Source Data
$SourceDataCSV = $SourceDataPath
$SourceDistributionGroups = Import-CSV $SourceDataCSV

#Get Destination Data
$DestinationDistributionGroups = Get-Destination-DistributionGroups

#Discover existing groups
$results = Compare-DistributionGroups -SourceDistributionGroups $SourceDistributionGroups -DestinationDistributionGroups $DestinationDistributionGroups 

#Create groups that do not exist
$nonExistantGroupsUnique = $results.UnMatched | sort GroupName -Unique
if($nonExistantGroupsUnique){
    foreach($group in $nonExistantGroupsUnique){
        try {
            $DistributionGroup = New-DistributionGroup -Name $group.GroupName -ErrorAction Stop
            $global:createdDistributionGroup += $DistributionGroup
        } catch {
            Write-Host "Failed to create DL."
            $global:failedToCreateDistributionGroup += $group
        }
    }
} else {
    Write-Host "No groups to create." -ForegroundColor Yellow
}

Add-Remove-Users-To-Groups -SourceDistributionGroups $SourceDistributionGroups

$global:usersNotFound
$global:addedToDistributionGroup
$global:createdDistributionGroup
$global:failedToCreateDistributionGroup