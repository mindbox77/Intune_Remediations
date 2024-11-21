
# Check for registry key to determine if New Outlook toggle is hidden
$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$RegistryName = "HideNewOutlookToggle"
$RegistryKeyExists = (Get-ItemProperty -Path $RegistryPath -Name $RegistryName -ErrorAction SilentlyContinue)

if ($RegistryKeyExists -eq $null -or $RegistryKeyExists.$RegistryName -ne 1) {
    Write-Output "New Outlook toggle is not hidden."
    exit 1
}

# Check if New Outlook appx package is installed
$OutlookAppxPackage = Get-AppxPackage -Name "Microsoft.OutlookForWindows" -AllUsers -ErrorAction SilentlyContinue
if ($OutlookAppxPackage) {
    Write-Output "New Outlook appx package is installed."
    exit 1
}

# Check if Office 365 New Outlook is installed (assuming it is deployed as Click-to-Run)
$ClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
if (Test-Path $ClickToRunPath) {
    $OfficeVersion = (& "$ClickToRunPath" /config) | Select-String "NewOutlook"
    if ($OfficeVersion) {
        Write-Output "Office 365 New Outlook is installed."
        exit 1
    }
}

# Additional registry key to check if New Outlook preview experience is disabled
$RegistryPathPreview = "HKCU:\Software\Microsoft\Office\Outlook\Addins"
$RegistryNamePreview = "NewOutlookEnabled"
$RegistryKeyPreviewExists = (Get-ItemProperty -Path $RegistryPathPreview -Name $RegistryNamePreview -ErrorAction SilentlyContinue)

if ($RegistryKeyPreviewExists -eq $null -or $RegistryKeyPreviewExists.$RegistryNamePreview -ne 0) {
    Write-Output "New Outlook preview experience is not disabled."
    exit 1
}

# If all checks pass, detection is successful
Write-Output "New Outlook for Windows and Office 365 New Outlook are not installed or are correctly blocked."
exit 0
