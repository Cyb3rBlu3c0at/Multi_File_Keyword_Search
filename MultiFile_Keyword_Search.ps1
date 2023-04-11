<#
Description: PowerShell script to automate the search of Multiple CSV files.

Author: Mike Dunn

Creation Date: 4/11/2023

Version: 1

NOTE: Requires a keyword text file.
#>

Add-Type -AssemblyName System.Windows.Forms
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select Keyword Text File"
$fileDialog.Filter = "TXT files (*.txt)|*.txt"

if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $KeywordFilePath = $fileDialog.FileName
    Write-Host "Selected file: $KeywordFilePath"
} else {
    Write-Host "File selection canceled"
    break
}

Add-Type -AssemblyName System.Windows.Forms
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select CSV files"
$fileDialog.Filter = "CSV files (*.csv)|*.csv"
$fileDialog.Multiselect = $true

if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVFilesPath = $fileDialog.FileNames
    foreach($FileName in $CSVFilesPath){
    Write-Host "Selected file: $FileName"
    }
} else {
    Write-Host "File selection canceled"
    break
}

$keywordList = New-Object System.Collections.ArrayList
Get-Content $KeywordFilePath | ForEach-Object { $keywordList.Add($_) | Out-Null }

Write-Host "Importing CSV Files, please wait..."
$csv = Import-Csv $CSVFilesPath

$totalRows = $csv.Count
$currentRow = 0
$percentComplete = 0
Write-Progress -Activity "Searching for matches" -Status "Processing row $currentRow of $totalRows" -PercentComplete $percentComplete


foreach ($row in $csv) {
    foreach ($keyword in $keywordList) {
        if ($row -match $keyword) {
            $row | Export-Csv -Path "$env:UserProfile\Desktop\Results.csv" -Append -Force -NoTypeInformation
            break
        }
    }
    $currentRow++
    $percentComplete = [int]($currentRow / $totalRows * 100)
    Write-Progress -Activity "Searching for matches" -Status "Processing row $currentRow of $totalRows" -PercentComplete $percentComplete
}

Write-Progress -Activity "Searching for matches" -Status "Complete" -PercentComplete 100