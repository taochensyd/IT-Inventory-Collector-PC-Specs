# Get CPU details
$cpu = Get-WmiObject -Class Win32_Processor

# Get RAM details
$ramModules = Get-WmiObject -Class Win32_PhysicalMemory
$ramSize = $ramModules | Measure-Object -Property Capacity -Sum
$ramFrequency = $ramModules | ForEach-Object { $_.Speed } 
$dimmSlotsUsed = $ramModules.Count

# Get the total DIMM slots
$dimmSlotsTotal = (Get-WmiObject -Class Win32_PhysicalMemoryArray).MemoryDevices

# Get the current logged in username
$currentUsername = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Get the IP Address that matches the pattern 192.168.0.X
$ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -like "192.168.*.*" }).IPAddress

# Fetching disk details
$logicalDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"
$physicalDrives = Get-WmiObject -Class Win32_DiskDrive

# Displaying information
Write-Output "---------- System Information ----------"
Write-Output ("CPU Name: " + $cpu.Name)
Write-Output ("CPU Cores: " + $cpu.NumberOfCores)
Write-Output ("CPU Threads: " + $cpu.ThreadCount)
Write-Output ("CPU Frequency(GHz): " + [math]::Round($cpu.MaxClockSpeed / 1000, 2))
Write-Output ("RAM Size(GB): " + [math]::Round($ramSize.Sum / 1GB, 2))
Write-Output ("DIMM Slots Used: " + $dimmSlotsUsed + "/" + $dimmSlotsTotal)
Write-Output ("RAM Frequency(MHz): " + ($ramFrequency -join " "))
Write-Output ("Logged In User: " + $currentUsername)
Write-Output ("IP Address: " + $ipAddress)

$logicalDisks | ForEach-Object {
    $relatedPhysicalDrive = $physicalDrives | Where-Object { $_.DeviceID -eq $_.DeviceID }
    $diskType = if (3 -in $relatedPhysicalDrive.Capabilities) { "SSD" } else { "HDD" }
    
    Write-Output ("Drive Letter: " + $_.DeviceID + " Total Size(GB): " + [math]::Round($_.Size / 1GB, 2) + 
                  " Used Size(GB): " + [math]::Round(($_.Size - $_.FreeSpace) / 1GB, 2) + " Type: " + $diskType)
}

Write-Output "---------------------------------------"


# Split the $currentUsername to extract PC Name and User Name
$splitUsername = $currentUsername -split "\\"
$pcName = $splitUsername[0]
$loggedInUser = $splitUsername[1]

# Create a JSON object with the gathered data
$jsonData = @{
    "cpuName"          = $cpu.Name
    "cpuCores"         = $cpu.NumberOfCores
    "cpuThreads"       = $cpu.ThreadCount
    "cpuFrequencyGHz"  = [math]::Round($cpu.MaxClockSpeed / 1000, 2)
    "ramSizeGB"        = [math]::Round($ramSize.Sum / 1GB, 2)
    "ramFrequencyMHz"  = ($ramFrequency -join " ")
    "dimmSlotsUsed"    = "$dimmSlotsUsed/$dimmSlotsTotal"
    "pcName"           = $pcName
    "loggedInUser"     = $loggedInUser
    "ipAddress"        = $ipAddress
    "disks"            = $logicalDisks | ForEach-Object {
        $relatedPhysicalDrive = $physicalDrives | Where-Object { $_.DeviceID -eq $_.DeviceID }
        $diskType = if (3 -in $relatedPhysicalDrive.Capabilities) { "SSD" } else { "HDD" }

        @{
            "driveLetter" = $_.DeviceID
            "totalSizeGB" = [math]::Round($_.Size / 1GB, 2)
            "usedSizeGB"  = [math]::Round(($_.Size - $_.FreeSpace) / 1GB, 2)
            "diskType"    = $diskType
        }
    }
} | ConvertTo-Json


# Send POST request to the specified API
Invoke-RestMethod -Method Post -Uri "http://localhost:3000/api/v1/pcinfo" -Body $jsonData -ContentType "application/json"
