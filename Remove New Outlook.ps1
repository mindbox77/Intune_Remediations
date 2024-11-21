# Uninstalling the New Outlook for Windows (Preview) and O365 New Outlook
echo "Removing New Outlook and blocking reinstallation..."

# Define the registry key to block reinstallation
$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$RegistryName = "HideNewOutlookToggle"

# Ensure the path exists
if (-not (Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force | Out-Null
}

# Set registry key to hide New Outlook toggle and prevent reinstallation
Set-ItemProperty -Path $RegistryPath -Name $RegistryName -Value 1 -Type DWord

# Remove the New Outlook appx package if installed
# Requires elevated privileges to run
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as an administrator to remove the New Outlook appx package. Please run the script with elevated privileges."
    exit
}

Get-AppxPackage -Name "Microsoft.OutlookForWindows" -AllUsers | Remove-AppxPackage -AllUsers

# Remove Office 365 New Outlook client if installed (assuming it is deployed as Click-to-Run)
$ClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
if (Test-Path $ClickToRunPath) {
    Start-Process -FilePath $ClickToRunPath -ArgumentList "/update user updatetoversion=previous" -Wait
}

# Additional registry key to disable the Outlook preview experience
$RegistryPathPreview = "HKCU:\Software\Microsoft\Office\Outlook\Addins"
$RegistryNamePreview = "NewOutlookEnabled"

# Ensure the path exists
if (-not (Test-Path $RegistryPathPreview)) {
    New-Item -Path $RegistryPathPreview -Force | Out-Null
}

# Set registry key to disable New Outlook preview experience
Set-ItemProperty -Path $RegistryPathPreview -Name $RegistryNamePreview -Value 0 -Type DWord

# Confirmation
Write-Output "New Outlook for Windows and Office 365 New Outlook have been removed and blocked from reinstalling."

# Exit with code 0 to indicate successful remediation
exit 0
