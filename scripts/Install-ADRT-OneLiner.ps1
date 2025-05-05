# Install-ADRT-OneLiner.ps1
# Este comando permite instalar o ADRT com uma única linha de PowerShell

# Para instalar o ADRT, basta executar o seguinte comando no PowerShell como administrador:
# iex (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/lobios/adrt-modern/main/Setup-ADRT-Online.ps1')

$scriptUrl = 'https://raw.githubusercontent.com/lobios/adrt-modern/main/Setup-ADRT-Online.ps1'

Write-Host @"
╔═══════════════════════════════════════════════════════════════╗
║                ADRT - One-Line Installer                      ║
╚═══════════════════════════════════════════════════════════════╝

Para instalar o ADRT Moderno com uma única linha de comando, 
copie e cole o seguinte comando no PowerShell como administrador:

iex (New-Object Net.WebClient).DownloadString('$scriptUrl')

Este comando irá baixar e executar o instalador completo diretamente do GitHub,
configurando todos os arquivos necessários para o funcionamento do ADRT.

Visite https://github.com/lobios/adrt-modern para mais informações.
"@ -ForegroundColor Cyan

# Se este script for executado diretamente, podemos iniciar a instalação automaticamente
$startInstall = Read-Host "Deseja iniciar a instalação agora? (S/N)"
if ($startInstall -eq "S" -or $startInstall -eq "s") {
    try {
        Write-Host "Iniciando instalação do ADRT Moderno..." -ForegroundColor Green
        Invoke-Expression (New-Object Net.WebClient).DownloadString($scriptUrl)
    }
    catch {
        Write-Host "Erro ao iniciar a instalação: $_" -ForegroundColor Red
        Write-Host @"
Por favor, execute manualmente o comando:
iex (New-Object Net.WebClient).DownloadString('$scriptUrl')
"@ -ForegroundColor Yellow
    }
}