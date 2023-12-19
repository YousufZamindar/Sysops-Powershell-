# Define the function to gather server configuration information
function Get-ServerConfiguration {
    param (
        [string]$ServerName
    )
 
    $systemInfo = @{
        'ServerName' = $ServerName
        'OperatingSystem' = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
        'OSVersion' = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
        'Manufacturer' = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
        'Model' = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
        'Processor' = (Get-CimInstance -ClassName Win32_Processor).Name
        'InstalledSoftware' = Get-WmiObject -Class Win32_Product | Select-Object -Property Name, Version, Vendor
        'RolesAndFeatures' = Get-WindowsFeature | Where-Object { $_.InstallState -eq 'Installed' } | Select-Object -Property Name, DisplayName, Description
    }
 
    return New-Object -TypeName PSObject -Property $systemInfo
}
 
# Read the list of server names from a text file
$ServerList = Get-Content -Path "C:\Path\To\Your\File\servers.txt"
 
# Create an array to hold the results
$Results = @()
 
# Iterate through each server name in the list and gather configuration information
foreach ($Server in $ServerList) {
    $Result = Invoke-Command -ComputerName $Server -ScriptBlock {
        Get-ServerConfiguration -ServerName $using:Server
    }
 
    $Results += $Result
}
 
# Convert the results to a formatted Excel file
$ExcelOutput = $Results | Select-Object ServerName, OperatingSystem, OSVersion, Manufacturer, Model, Processor, InstalledSoftware, RolesAndFeatures
$ExcelOutput | Export-Excel -Path "C:\Path\To\Your\File\Server_Configuration.xlsx" -AutoSize -FreezeTopRow -AutoFilter