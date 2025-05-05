# Uninstall-ADRT.ps1
$confirmation = Read-Host "Are you sure you want to remove all ADRT reports? (Y/N)"
if ($confirmation -eq 'Y') {
    # Remove report directories
    Remove-Item -Path ".\ad-reports" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "ADRT reports have been removed." -ForegroundColor Green
}