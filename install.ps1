# install.ps1
# Adds the sideload registry entry for Excel Office Add-in

$manifestPath = "D:\Downloads\manifest.xml"

# Create registry key
New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Force | Out-Null

# Set manifest path
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" `
  -Name "Manifest" `
  -Value $manifestPath

# Show success popup
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("✅ Add-in sideloaded successfully from:`n$manifestPath", "Excel Add-in Installer", 'OK', 'Information')
