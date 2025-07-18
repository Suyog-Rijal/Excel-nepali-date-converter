# uninstall.ps1
# Removes the sideloaded Office Add-in registry entry

$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

Add-Type -AssemblyName System.Windows.Forms

if (Test-Path $regPath) {
    Remove-Item -Path $regPath -Recurse -Force
    [System.Windows.Forms.MessageBox]::Show("🧹 Sideloaded Excel Add-in removed successfully.", "Excel Add-in Uninstaller", 'OK', 'Information')
} else {
    [System.Windows.Forms.MessageBox]::Show("ℹ️ No sideloaded add-in found to remove.", "Excel Add-in Uninstaller", 'OK', 'Warning')
}
