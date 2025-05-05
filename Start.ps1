#Requires -RunAsAdministrator

# Verificar vers�o do PowerShell
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Este script requer o PowerShell 7 ou superior." -ForegroundColor Red
    Write-Host "Voc� est� usando o PowerShell $($PSVersionTable.PSVersion)."
    Write-Host ""
    Write-Host "Para utilizar todas as funcionalidades do ADRT, � necess�rio instalar o PowerShell 7." -ForegroundColor Yellow
    Write-Host "Link para download: https://aka.ms/powershell-release?tag=stable" -ForegroundColor Cyan
    Write-Host ""
    $choice = Read-Host "Deseja abrir o link para baixar o PowerShell 7? (S/N)"
    
    if ($choice -eq "S" -or $choice -eq "s") {
        Start-Process "https://aka.ms/powershell-release?tag=stable"
    }
    
    Write-Host "Por favor, instale o PowerShell 7 e execute novamente este script." -ForegroundColor Green
    Read-Host "Pressione Enter para sair"
    exit
}

# Se chegou aqui, est� usando PowerShell 7+
Write-Host "PowerShell 7 detectado. Iniciando ADRT..." -ForegroundColor Green
Write-Host "Vers�o: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan

# Aguardar um momento e continuar
Start-Sleep -Seconds 2

# Aqui voc� pode chamar seu script principal ou continuar com o c�digo
# Por exemplo:
& ".\Start-ADRT.ps1"

Write-Host "Script completo!" -ForegroundColor Green