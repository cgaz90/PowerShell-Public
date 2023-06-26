<#
    .SYNOPSIS
        A script that clears the Teams cache.

    .DESCRIPTION
        This script closes the Microsoft Teams application, then clears its cache by deleting the specific cache directories.

    .AUTHOR
        Garry Crothall 

    .DATE
        26/06/2023

    .USAGE
        Run the script in PowerShell. No arguments are needed.

    .DEPENDENCIES
        Microsoft Teams.

    .CHANGELOG
        26/06/2023 - Garry Crothall  - Improved error handling.
        26/06/2023 - Garry Crothall  - Added additional cache directories.
        26/06/2023 - Garry Crothall  - Fixed error handling output parser error.
#>

# Define process name
$processName = "Teams"

# Dynamically get the username
$username = [Environment]::UserName

# Define Teams cache directories
$cacheDirectories = @(
    "$env:APPDATA\Microsoft\Teams\Application Cache\Cache",
    "$env:APPDATA\Microsoft\Teams\Blob_storage",
    "$env:APPDATA\Microsoft\Teams\Cache",
    "$env:APPDATA\Microsoft\Teams\databases",
    "$env:APPDATA\Microsoft\Teams\GPUcache",
    "$env:APPDATA\Microsoft\Teams\IndexedDB",
    "$env:APPDATA\Microsoft\Teams\Local Storage",
    "$env:APPDATA\Microsoft\Teams\tmp"
)

# Check if Teams process is running and stop it
do {
    $process = Get-Process | Where-Object { $_.Name -eq $processName }

    if ($process) {
        # If the process is running, stop it
        Write-Output "Closing Microsoft Teams..."
        Stop-Process -Name $processName -Force
        # Give the process some time to close down
        Start-Sleep -Seconds 2
    }
} while ($process)

# Delete cache directories
Write-Output "Deleting cache files..."
foreach ($directory in $cacheDirectories) {
    try {
        if (Test-Path $directory) {
            Write-Output "Deleting $directory..."
            Remove-Item -Path $directory -Force -Recurse
        } else {
            Write-Warning "Directory not found: $directory"
        }
    } catch {
        Write-Error "Failed to delete ${directory}: $_"

    }
}

Write-Output "Cache cleanup complete."