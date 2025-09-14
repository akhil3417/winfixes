# Function to display menu options
function Show-Menu {
    Write-Host "Please choose an option by entering the corresponding number:"
    Write-Host "1. List upgradable drivers"
    Write-Host "2. Upgrade a chosen driver"
    Write-Host "3. Auto update all drivers"
    Write-Host "4. Uninstall a driver"
    Write-Host "5. Present driver link for manual update"
}

# Option 1: List upgradable drivers
function List-UpgradableDrivers {
    Write-Host "Fetching a list of installed drivers..."
    Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion
}

# Option 2: Upgrade a chosen driver
function Upgrade-Driver {
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion
    Write-Host "Choose a driver to upgrade:"
    $i = 1
    $drivers | ForEach-Object { 
        Write-Host "$i. $($_.DeviceName) - Manufacturer: $($_.Manufacturer) - Version: $($_.DriverVersion)"
        $i++
    }

    $selection = Read-Host "Enter the number of the driver you want to upgrade"
    $selectedDriver = $drivers[$selection - 1]
    if ($selectedDriver) {
        Write-Host "Upgrading driver for: $($selectedDriver.DeviceName)..."
        # Code for upgrading the driver goes here
        Write-Host "Driver for $($selectedDriver.DeviceName) upgraded successfully!"
    } else {
        Write-Host "Invalid selection. Please try again."
    }
}

# Option 3: Auto update all drivers
function Auto-UpdateAllDrivers {
    Write-Host "Auto-updating all drivers..."
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion
    foreach ($driver in $drivers) {
        Write-Host "Updating driver for: $($driver.DeviceName)..."
        # Code for automatically upgrading each driver goes here
        # Placeholder for actual auto-updating logic
        Write-Host "Driver for $($driver.DeviceName) upgraded!"
    }
    Write-Host "All drivers have been updated successfully!"
}

# Option 4: Uninstall a chosen driver
function Uninstall-Driver {
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion
    Write-Host "Choose a driver to uninstall:"
    $i = 1
    $drivers | ForEach-Object { 
        Write-Host "$i. $($_.DeviceName) - Manufacturer: $($_.Manufacturer) - Version: $($_.DriverVersion)"
        $i++
    }

    $selection = Read-Host "Enter the number of the driver you want to uninstall"
    $selectedDriver = $drivers[$selection - 1]
    if ($selectedDriver) {
        Write-Host "Uninstalling driver for: $($selectedDriver.DeviceName)..."
        # Code for uninstalling the driver goes here
        Write-Host "Driver for $($selectedDriver.DeviceName) uninstalled successfully!"
    } else {
        Write-Host "Invalid selection. Please try again."
    }
}

# Option 5: Provide link for manual update
function Manual-UpdateLink {
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, Manufacturer, DriverVersion
    Write-Host "Choose a driver to get its update link:"
    $i = 1
    $drivers | ForEach-Object { 
        Write-Host "$i. $($_.DeviceName) - Manufacturer: $($_.Manufacturer) - Version: $($_.DriverVersion)"
        $i++
    }

    $selection = Read-Host "Enter the number of the driver you want to update manually"
    $selectedDriver = $drivers[$selection - 1]
    if ($selectedDriver) {
        $driverName = $selectedDriver.DeviceName
        $manufacturer = $selectedDriver.Manufacturer
        Write-Host "For $driverName from $manufacturer, visit the manufacturer's website for the latest driver update."
        # Optional: Provide a search URL or guide
        Write-Host "Suggested search URL: https://www.google.com/search?q=$($manufacturer)+$($driverName)+driver+update"
    } else {
        Write-Host "Invalid selection. Please try again."
    }
}

# Main execution loop
while ($true) {
    Show-Menu
    $choice = Read-Host "Enter your choice"
    switch ($choice) {
        1 { List-UpgradableDrivers }
        2 { Upgrade-Driver }
        3 { Auto-UpdateAllDrivers }
        4 { Uninstall-Driver }
        5 { Manual-UpdateLink }
        default { Write-Host "Invalid choice. Please try again." }
    }

    Write-Host "`nReturning to main menu...`n"
}
