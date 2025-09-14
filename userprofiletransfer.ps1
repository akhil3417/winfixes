# Input from User
Write-Host 'Make sure the user has signed into both computers so their user profile will have been created.' -ForegroundColor Red 
$originalComputerName = Read-Host "What is the computer name you're moving from?"
$newComputerName = Read-Host 'What is the new computer name?'
$username = Read-Host 'What is the username?'
Write-Host 'What would you like to transfer?'
Write-Host '1. Desktop' -ForegroundColor Blue
Write-Host '2. Documents' -ForegroundColor Yellow
Write-Host '3. Downloads' -ForegroundColor Cyan
Write-Host '4. Browser Data' -ForegroundColor Magenta
Write-Host '5. Everything' -ForegroundColor Red
$selection = Read-Host 'Enter 1-5.'

# Copy From Here
$oldDesktop = "\\$originalComputerName\c$\Users\$username\Desktop\*"
$oldDownloads = "\\$originalComputerName\c$\Users\$username\Downloads\*"
$oldDocuments = "\\$originalComputerName\c$\Users\$username\Documents\*"
$oldChromeData = "\\$originalComputerName\c$\Users\$username\AppData\Local\Google\Chrome\User Data\Default\*"
$oldEdgeData = "\\$originalComputerName\c$\Users\$username\AppData\Local\Microsoft\Edge\User Data\Default\*"
$oldFirefoxData = "\\$originalComputerName\c$\Users\$username\AppData\Roaming\Mozilla\Firefox\Profiles\*"

# Copy To Here
$newDesktop = "\\$newComputerName\c$\Users\$username\Desktop\"
$newDownloads = "\\$newComputerName\c$\Users\$username\Downloads\"
$newDocuments = "\\$newComputerName\c$\Users\$username\Documents\"
$newChromeData = "\\$newComputerName\c$\Users\$username\AppData\Local\Google\Chrome\User Data\Default\"
$newEdgeData = "\\$newComputerName\c$\Users\$username\AppData\Local\Microsoft\Edge\User Data\Default\"
$newFirefoxData = "\\$newComputerName\c$\Users\$username\AppData\Roaming\Mozilla\Firefox\Profiles\"

# Copy Operations & Selection
$copyOperations = @()

switch ($selection) {
    '1' { 
        $copyOperations += @{
            Source = $oldDesktop
            Destination = $newDesktop
            Description = "Copying Desktop Files"
        }
    }
    '2' { 
        $copyOperations += @{
            Source = $oldDocuments
            Destination = $newDocuments
            Description = "Copying Documents"
        }
    }
    '3' { 
        $copyOperations += @{
            Source = $oldDownloads
            Destination = $newDownloads
            Description = "Copying Downloads"
        }
    }
    '4' { 
        $copyOperations += @{
            Source = $oldChromeData
            Destination = $newChromeData 
            Description = "Copying Chrome Data"
        }
        $copyOperations += @{
            Source = $oldEdgeData
            Destination = $newEdgeData 
            Description = "Copying Edge Data"
        }
        $copyOperations += @{
            Source = $oldFirefoxData
            Destination = $newFirefoxData 
            Description = "Copying Firefox Data"
        }
    }
    '5' { 
        $copyOperations += @{
            Source = $oldDesktop
            Destination = $newDesktop
            Description = "Copying Desktop Files"
        }
        $copyOperations += @{
            Source = $oldDocuments
            Destination = $newDocuments
            Description = "Copying Documents"
        }
        $copyOperations += @{
            Source = $oldDownloads
            Destination = $newDownloads
            Description = "Copying Downloads"
        }
        $copyOperations += @{
            Source = $oldChromeData
            Destination = $newChromeData 
            Description = "Copying Chrome Data"
        }
        $copyOperations += @{
            Source = $oldEdgeData
            Destination = $newEdgeData 
            Description = "Copying Edge Data"
        }
        $copyOperations += @{
            Source = $oldFirefoxData
            Destination = $newFirefoxData 
            Description = "Copying Firefox Data"
        }
    }
    default { 
        Write-Host "Invalid selection. Exiting." -ForegroundColor Red
        exit
    }
}


# Total number of operations for progress tracking
$totalOperations = $copyOperations.Count

# Perform copy operations with progress
foreach ($operation in $copyOperations) {
    $currentOperation = $copyOperations.IndexOf($operation) + 1
    Write-Progress -Activity "Copying Files" `
                   -Status $operation.Description `
                   -PercentComplete (($currentOperation / $totalOperations) * 100)
    
    try {
        Copy-Item -Path $operation.Source -Destination $operation.Destination -Recurse -ErrorAction SilentlyContinue
        Write-Host "Successfully copied $($operation.Description)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error copying $($operation.Description): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Complete the progress bar
Write-Progress -Activity "Copying Files" -Completed
Write-Host "File copy process completed." -ForegroundColor Cyan

# Keep Window Open
Read-Host "Press Enter to exit"
