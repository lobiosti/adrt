# ==========================================================
# Validação e instalação do PowerShell 7 (se necessário)
# ==========================================================
function Test-PwshInstalled {
    return (Get-Command "pwsh" -ErrorAction SilentlyContinue) -ne $null
}

function Install-Pwsh {
    Write-Host "Baixando e instalando o PowerShell 7..." -ForegroundColor Yellow
    $installerPath = "$env:TEMP\PowerShell-7-x64.msi"
    $downloadUrl = "https://github.com/PowerShell/PowerShell/releases/latest/download/PowerShell-7.4.2-win-x64.msi"

    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

    Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait

    Remove-Item $installerPath -Force

    # Atualiza PATH da sessão, se necessário
    $env:PATH += ";$env:ProgramFiles\PowerShell\7\"
}

if (-not (Test-PwshInstalled)) {
    Write-Warning "O PowerShell 7 (pwsh.exe) não está instalado neste sistema."
    $installNow = Read-Host "Deseja instalar agora? (S/N)"

    if ($installNow.ToUpper() -eq "S") {
        Install-Pwsh

        if (-not (Test-PwshInstalled)) {
            Write-Host "A instalação falhou ou o pwsh ainda não está acessível. Encerrando..." -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Instalação do PowerShell 7 cancelada. Encerrando..." -ForegroundColor Red
        exit 1
    }
}

# ==========================================================
# Script de instalação seguro para ADRT
# ==========================================================
Write-Host "Iniciando instalação segura do ADRT..." -ForegroundColor Cyan

# 1. Definir o diretório de instalação
$installDir = "$env:USERPROFILE\ADRT"
Write-Host "Diretório de instalação: $installDir" -ForegroundColor Yellow

# 2. Criar diretório se não existir
if (-not (Test-Path -Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    Write-Host "Diretório de instalação criado" -ForegroundColor Green
}

# 3. Baixar o arquivo ZIP do repositório
$zipUrl = "https://github.com/lobiosti/adrt/archive/refs/heads/main.zip"
$zipFile = "$env:TEMP\adrt.zip"
Write-Host "Baixando ADRT..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile
Write-Host "Download concluído" -ForegroundColor Green

# 4. Extrair o conteúdo
Write-Host "Extraindo arquivos..." -ForegroundColor Yellow
Expand-Archive -Path $zipFile -DestinationPath $env:TEMP -Force
Copy-Item -Path "$env:TEMP\adrt-main\*" -Destination $installDir -Recurse -Force
Write-Host "Extração concluída" -ForegroundColor Green

# 5. Limpar arquivos temporários
Remove-Item -Path $zipFile -Force
Remove-Item -Path "$env:TEMP\adrt-main" -Recurse -Force

# 6. Criar atalhos (opcional)
$desktopShortcut = "$env:USERPROFILE\Desktop\ADRT.lnk"
$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($desktopShortcut)
$shortcut.TargetPath = "$installDir\Start-ADRT.ps1"
$shortcut.WorkingDirectory = $installDir
$shortcut.Save()

Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                      INSTALAÇÃO CONCLUÍDA                     ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "O ADRT foi instalado em: $installDir" -ForegroundColor Green
Write-Host "Um atalho foi criado na sua área de trabalho." -ForegroundColor Green
Write-Host ""
Write-Host "Para iniciar o ADRT, execute o atalho ou navegue até o diretório"
Write-Host "de instalação e execute o arquivo Start-ADRT.ps1 como administrador."
Write-Host ""
Write-Host "Desenvolvido por Lobios Segurança • Tecnologia • Inovação"

# Perguntar se deseja iniciar o ADRT agora
$startNow = Read-Host "Deseja iniciar o ADRT agora? (S/N)"

if ($startNow.ToUpper() -eq "S") {
    Write-Host "Iniciando o ADRT com PowerShell 7..." -ForegroundColor Cyan
    Start-Process pwsh -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$installDir\Start-ADRT.ps1`""
}
