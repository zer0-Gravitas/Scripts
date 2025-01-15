<#
.SYNOPSIS
    This script gathers a list of files in a directory and logs its metadata to a .csv file. This script will also check xlsx files to see if they contain links.

    This is compatible with PowerShell 5.x

.DESCRIPTION
    The script will read a specified directory and gather the metadata it will then create a csv file with the following attributes:
        String - FullName - FilePath.
        String - Name - Name of the file.
        Int - SizeMB - The size of the file in MB.
        String - Owner - The owner of the file.
        DateTime - LastAccessTime - The last access data and time of the file.
        String - Extension - The file extension.
        Bool - Active - Has the file has been accessed in $ExpirationMonths.
        Bool - Unwanted = Is the file of a type that should be kept $UnwantedExtensions.
        Bool - ContainsLinks = If the file is xlsx, and is active does it contain and link.

.PARAMETER DirectoryPath
    The path of the directory to be scanned.

.PARAMETER OutputFilePath
    The path and name of the csv that will contain the generated data.

.PARAMETER ExpirationMonths
    How many months until a file is considered no longer active. Default 12.

.PARAMETER UnwantedExtensions
    A list of unwanted extensions, should be specified comma seperated.
    -UnwantedExtensions ".tmp", ".log", ".bak"

.PARAMETER MaxThreads
    The number of cores to use in the run, default 2.

.EXAMPLE
    .\data_analysis.ps1 `
        -DirectoryPath "D:\Data" `
        -OutputFilePath c:\temp\data_export.csv `
        -ExpirationMonths 12 `
        -UnwantedExtensions ".tmp", ".log", ".bak", ".exe" `
        -MaxThreads 2

.NOTES
    Version:        1.0
    Author:         Brian Jones
    Creation Date:  2025-01-15
    Last Modified:  2025-01-15
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$DirectoryPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputFilePath,

    [Parameter(Mandatory = $false)]
    [int]$ExpirationMonths = 12,

    [Parameter(Mandatory = $false)]
    [string[]]$UnwantedExtensions = @(".tmp", ".log", ".bak"),

    [Parameter(Mandatory = $false)]
    [int]$MaxThreads = 4
)

if (-not (Test-Path -Path $DirectoryPath)) {
    throw "The directory path '$DirectoryPath' does not exist."
}

$Files = Get-ChildItem -Path $DirectoryPath -Recurse -File

# Create runspace pool
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$RunspacePool.Open()

# Create collection for results
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Results = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))

# Define the script block to run in runspaces
$ScriptBlock = {
    param ($File, $ExpirationMonths, $UnwantedExtensions, $SyncHash)

    try {
        $CurrentDate = Get-Date
        $IsActive = if (($CurrentDate - $File.LastAccessTime).Days -lt ($ExpirationMonths * 30)) { $true } else { $false }
        $IsUnwanted = if ($UnwantedExtensions -contains $File.Extension.ToLower()) { $true } else { $false }
        $SizeMB = [math]::Round($File.Length / 1MB, 2)
        $ContainsLinks = $false
        if ($IsActive -and $_.Extension -eq ".xlsx") {
            try {
                $Excel = New-Object -ComObject Excel.Application
                $Workbook = $Excel.Workbooks.Open($_.FullName, [Type]::Missing, $true)
                foreach ($Worksheet in $Workbook.Worksheets) {
                    if ($Worksheet.Hyperlinks.Count -gt 0) {
                        $ContainsLinks = $true
                        break
                    }
                }
                $Workbook.Close($false)
                $Excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
            } catch {
                Write-Host "Error checking links in file: $($_.FullName). Error: $_" -ForegroundColor Yellow
            }
        }   
        $Result = [PSCustomObject]@{
            FullName      = $File.FullName
            Name          = $File.Name
            SizeMB        = $SizeMB
            Owner         = (Get-Acl $File.FullName).Owner
            LastAccessTime = $File.LastAccessTime
            Extension     = $File.Extension
            Active        = $IsActive
            Unwanted      = $IsUnwanted
            ContainsLinks = $ContainsLinks
        }

        $SyncHash.Results.Add($Result) | Out-Null
    } catch {
        Write-Host "Error processing file: $($File.FullName). Error: $_" -ForegroundColor Red
    }
}

foreach ($File in $Files) {
    $PowerShell = [powershell]::Create().AddScript($ScriptBlock).AddArgument($File).AddArgument($ExpirationMonths).AddArgument($UnwantedExtensions).AddArgument($SyncHash)
    $PowerShell.RunspacePool = $RunspacePool
    $Runspaces += [PSCustomObject]@{
        PowerShell = $PowerShell
        Handle     = $PowerShell.BeginInvoke()
    }
}

foreach ($Runspace in $Runspaces) {
    $Runspace.PowerShell.EndInvoke($Runspace.Handle)
    $Runspace.PowerShell.Dispose()
}

$RunspacePool.Close()
$RunspacePool.Dispose()

# Deduplicate
$DeduplicatedFiles = $SyncHash.Results | Group-Object -Property Owner, Name, SizeMB | ForEach-Object {
    $_.Group | Select-Object -First 1
}

$DeduplicatedFiles | Export-Csv -Path $OutputFilePath -NoTypeInformation -Force

Write-Host "Processed file metadata has been exported to: $OutputFilePath" -ForegroundColor Green