# Setup-ADRT-Online.ps1
# Script interativo para configuração do ADRT Moderno
# Baixa automaticamente os arquivos necessários do GitHub

<#
.SYNOPSIS
    Script de instalação do ADRT Moderno com recursos de download automático
.DESCRIPTION
    Este script interativo configura o ADRT Moderno, baixando os arquivos
    necessários do GitHub e criando a estrutura de diretórios completa.
.NOTES
    Autor: Lobios Segurança • Tecnologia • Inovação
    Versão: 2.0 (Instalação Online)
#>

# Definir codificação para garantir acentuação correta
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding'] = 'UTF8'

# URLs e informações de repositório
$repoOwner = "lobios"
$repoName = "adrt-modern"
$repoBranch = "main"
$baseUrl = "https://raw.githubusercontent.com/$repoOwner/$repoName/$repoBranch"

# Lista de arquivos principais para download
$coreFiles = @(
    "index-modern.html",
    "README.md",
    "Start-ADRT.ps1",
    "setup-ADRT.ps1",
    "Uninstall-ADRT.ps1",
    "analise-completa.ps1",
    "ad-users-modern.ps1",
    "ad-admins-modern.ps1",
    "ad-enterprise-admins-modern.ps1",
    "ad-disabled-modern.ps1",
    "ad-lastlogon-modern.ps1",
    "ad-neverexpires-modern.ps1",
    "ad-groups-modern.ps1",
    "ad-membergroups-modern.ps1",
    "ad-ous-modern.ps1",
    "ad-computers-modern.ps1",
    "ad-servers-modern.ps1",
    "ad-dcs-modern.ps1",
    "ad-gpos-modern.ps1",
    "ad-inventory-modern.ps1"
)

# Lista de diretórios a serem criados
$directories = @(
    "ad-reports\ad-users",
    "ad-reports\ad-admins",
    "ad-reports\ad-enterprise-admins",
    "ad-reports\ad-disabled",
    "ad-reports\ad-lastlogon",
    "ad-reports\ad-neverexpires",
    "ad-reports\ad-groups",
    "ad-reports\ad-membergroups",
    "ad-reports\ad-ous",
    "ad-reports\ad-computers",
    "ad-reports\ad-servers",
    "ad-reports\ad-dcs",
    "ad-reports\ad-gpos",
    "ad-reports\ad-inventory",
    "ad-reports\ad-all",
    "ad-reports\ad-analysis",
    "web\img",
    "modules",
    "templates",
    "config"
)

# Recursos adicionais
$resources = @(
    @{
        URL = "$baseUrl/web/img/lobios-logo.png"
        Path = "web\img\lobios-logo.png"
        Type = "Image"
    },
    @{
        URL = "$baseUrl/web/img/lobios-logo-small.png"
        Path = "web\img\lobios-logo-small.png"
        Type = "Image"
    },
    @{
        URL = "$baseUrl/templates/modern-template.html"
        Path = "templates\modern-template.html"
        Type = "Template"
    },
    @{
        URL = "$baseUrl/modules/ADRT-Modern.psm1"
        Path = "modules\ADRT-Modern.psm1"
        Type = "Module"
    },
    @{
        URL = "$baseUrl/config/config.txt"
        Path = "config\config.txt"
        Type = "Config"
    }
)

# Banner
function Show-Banner {
    Clear-Host
    Write-Host @"

╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║      ██╗      ██████╗ ██████╗ ██╗ ██████╗ ███████╗           ║
║      ██║     ██╔═══██╗██╔══██╗██║██╔═══██╗██╔════╝           ║
║      ██║     ██║   ██║██████╔╝██║██║   ██║███████╗           ║
║      ██║     ██║   ██║██╔══██╗██║██║   ██║╚════██║           ║
║      ███████╗╚██████╔╝██████╔╝██║╚██████╔╝███████║           ║
║      ╚══════╝ ╚═════╝ ╚═════╝ ╚═╝ ╚═════╝ ╚══════╝           ║
║                                                               ║
║      ADRT - Active Directory Report Tool (Instalação)         ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Magenta

    Write-Host "Bem-vindo ao instalador do ADRT Moderno!" -ForegroundColor Cyan
    Write-Host "Este script irá configurar o ADRT, baixando todos os arquivos necessários." -ForegroundColor Cyan
    Write-Host ""
}

# Função para verificar requisitos
function Test-Requirements {
    Write-Host "Verificando requisitos do sistema..." -ForegroundColor Cyan
    
    # Verificar PowerShell versão 5.1 ou superior
    $psVersion = $PSVersionTable.PSVersion
    Write-Host "PowerShell versão: $($psVersion.Major).$($psVersion.Minor)" -ForegroundColor Gray
    
    if ($psVersion.Major -lt 5) {
        Write-Host "⚠️ AVISO: PowerShell 5.1 ou superior é recomendado para melhor compatibilidade" -ForegroundColor Yellow
    } else {
        Write-Host "✅ PowerShell versão compatível" -ForegroundColor Green
    }
    
    # Verificar módulo ActiveDirectory
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "✅ Módulo ActiveDirectory está instalado e disponível" -ForegroundColor Green
    } catch {
        Write-Host "❌ Módulo ActiveDirectory não encontrado" -ForegroundColor Red
        Write-Host "O ADRT requer o módulo ActiveDirectory, que faz parte das Ferramentas RSAT (Remote Server Administration Tools)" -ForegroundColor Yellow
        
        $installRSAT = Read-Host "Deseja tentar instalar as ferramentas RSAT agora? (S/N)"
        if ($installRSAT -eq "S" -or $installRSAT -eq "s") {
            try {
                # Tentar instalar RSAT-AD-PowerShell
                Write-Host "Instalando ferramentas RSAT para Active Directory..." -ForegroundColor Cyan
                Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
                Write-Host "✅ Ferramentas RSAT instaladas com sucesso" -ForegroundColor Green
                
                # Tentar importar novamente
                Import-Module ActiveDirectory -ErrorAction Stop
                Write-Host "✅ Módulo ActiveDirectory carregado com sucesso" -ForegroundColor Green
            } catch {
                Write-Host "❌ Falha ao instalar ferramentas RSAT: $_" -ForegroundColor Red
                Write-Host "Por favor, instale manualmente as ferramentas RSAT para Active Directory antes de continuar." -ForegroundColor Yellow
                Write-Host "No Windows 10/11, vá para Configurações > Aplicativos > Recursos opcionais > Adicionar um recurso" -ForegroundColor Yellow
                Write-Host "e procure por 'RSAT: Active Directory Domain Services and Lightweight Directory Services Tools'" -ForegroundColor Yellow
                return $false
            }
        } else {
            Write-Host "Por favor, instale manualmente as ferramentas RSAT para Active Directory antes de continuar." -ForegroundColor Yellow
            return $false
        }
    }
    
    # Verificar conexão com a internet
    Write-Host "Verificando conexão com a Internet..." -ForegroundColor Gray
    try {
        $webRequest = Invoke-WebRequest -Uri "https://github.com" -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        Write-Host "✅ Conexão com a Internet estabelecida" -ForegroundColor Green
    } catch {
        Write-Host "❌ Não foi possível conectar à Internet: $_" -ForegroundColor Red
        Write-Host "A instalação online requer acesso à Internet para baixar os arquivos necessários." -ForegroundColor Yellow
        return $false
    }
    
    # Verificar permissões de escrita no diretório atual
    try {
        $testFile = Join-Path -Path (Get-Location) -ChildPath "adrt-test.tmp"
        [System.IO.File]::WriteAllText($testFile, "Teste de permissão de escrita", [System.Text.Encoding]::UTF8)
        Remove-Item -Path $testFile -Force
        Write-Host "✅ Permissões de escrita verificadas" -ForegroundColor Green
    } catch {
        Write-Host "❌ Sem permissão de escrita no diretório atual: $_" -ForegroundColor Red
        Write-Host "Execute o PowerShell como administrador e tente novamente." -ForegroundColor Yellow
        return $false
    }
    
    return $true
}

# Função para criar estrutura de diretórios
function Create-DirectoryStructure {
    Write-Host ""
    Write-Host "Criando estrutura de diretórios..." -ForegroundColor Cyan
    
    $currentDir = (Get-Item -Path ".").FullName
    Write-Host "Diretório de instalação: $currentDir" -ForegroundColor Gray
    
    foreach ($dir in $directories) {
        $dirPath = Join-Path -Path $currentDir -ChildPath $dir
        if (-not (Test-Path -Path $dirPath)) {
            try {
                New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
                Write-Host "✅ Criado diretório: $dir" -ForegroundColor Green
            } catch {
                Write-Host "❌ Erro ao criar diretório $dir : $_" -ForegroundColor Red
            }
        } else {
            Write-Host "ℹ️ Diretório já existe: $dir" -ForegroundColor Gray
        }
    }
}

# Função para baixar arquivos do GitHub
function Download-Files {
    Write-Host ""
    Write-Host "Baixando arquivos do repositório..." -ForegroundColor Cyan
    
    # Download dos scripts principais
    foreach ($file in $coreFiles) {
        $url = "$baseUrl/$file"
        $destination = Join-Path -Path (Get-Location) -ChildPath $file
        
        try {
            Write-Host "Baixando $file..." -ForegroundColor Gray -NoNewline
            Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
            Write-Host " ✅" -ForegroundColor Green
        } catch {
            Write-Host " ❌ Falha: $_" -ForegroundColor Red
        }
    }
    
    # Download dos recursos adicionais
    foreach ($resource in $resources) {
        $destination = Join-Path -Path (Get-Location) -ChildPath $resource.Path
        
        try {
            Write-Host "Baixando $($resource.Type): $($resource.Path)..." -ForegroundColor Gray -NoNewline
            $destinationDir = Split-Path -Path $destination -Parent
            if (-not (Test-Path -Path $destinationDir)) {
                New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
            }
            
            Invoke-WebRequest -Uri $resource.URL -OutFile $destination -UseBasicParsing
            Write-Host " ✅" -ForegroundColor Green
        } catch {
            Write-Host " ❌ Falha: $_" -ForegroundColor Red
        }
    }
}

# Função para configurar arquivos
function Configure-Files {
    Write-Host ""
    Write-Host "Configurando arquivos..." -ForegroundColor Cyan
    
    # Corrigir encoding de arquivos HTML
    try {
        $htmlFiles = Get-ChildItem -Path (Join-Path -Path (Get-Location) -ChildPath "ad-reports") -Filter "*.html" -Recurse -ErrorAction SilentlyContinue
        
        if ($htmlFiles -and $htmlFiles.Count -gt 0) {
            Write-Host "Corrigindo codificação de $($htmlFiles.Count) arquivos HTML..." -ForegroundColor Gray
            
            foreach ($file in $htmlFiles) {
                try {
                    $content = Get-Content -Path $file.FullName -Raw -ErrorAction SilentlyContinue
                    if ($content) {
                        [System.IO.File]::WriteAllText($file.FullName, $content, [System.Text.UTF8Encoding]::new($false))
                    }
                } catch {
                    Write-Host "❌ Erro ao processar $($file.Name): $_" -ForegroundColor Red
                }
            }
        }
    } catch {
        Write-Host "❌ Erro ao buscar arquivos HTML: $_" -ForegroundColor Red
    }
    
    # Corrigir encoding de scripts PS1
    try {
        $psScripts = Get-ChildItem -Path (Get-Location) -Filter "*.ps1" -ErrorAction SilentlyContinue
        
        Write-Host "Corrigindo codificação de $($psScripts.Count) scripts PowerShell..." -ForegroundColor Gray
        
        foreach ($script in $psScripts) {
            try {
                $content = Get-Content -Path $script.FullName -Raw -ErrorAction SilentlyContinue
                if ($content) {
                    [System.IO.File]::WriteAllText($script.FullName, $content, [System.Text.UTF8Encoding]::new($false))
                }
            } catch {
                Write-Host "❌ Erro ao processar $($script.Name): $_" -ForegroundColor Red
            }
        }
    } catch {
        Write-Host "❌ Erro ao buscar scripts PS1: $_" -ForegroundColor Red
    }
    
    # Configurar arquivo de configuração
    $configPath = Join-Path -Path (Get-Location) -ChildPath "config\config.txt"
    if (Test-Path -Path $configPath) {
        Write-Host "Personalizando arquivo de configuração..." -ForegroundColor Gray
        
        # Obter informações de domínio
        try {
            $domainInfo = Get-ADDomain
            $domainName = $domainInfo.DNSRoot
            
            # Personalizar arquivo de configuração
            $configContent = Get-Content -Path $configPath -Raw
            $configContent = $configContent -replace "{DOMINIO}", $domainName
            
            # Solicitar nome da empresa
            $companyName = Read-Host "Digite o nome da sua empresa ou organização"
            if ([string]::IsNullOrEmpty($companyName)) {
                $companyName = "Lobios Technology"
            }
            $configContent = $configContent -replace "{EMPRESA}", $companyName
            
            # Solicitar nome do responsável
            $ownerName = Read-Host "Digite o nome do responsável pela geração dos relatórios"
            if ([string]::IsNullOrEmpty($ownerName)) {
                $ownerName = "Administrador"
            }
            $configContent = $configContent -replace "{RESPONSAVEL}", $ownerName
            
            # Salvar configurações
            [System.IO.File]::WriteAllText($configPath, $configContent, [System.Text.UTF8Encoding]::new($false))
            Write-Host "✅ Arquivo de configuração personalizado" -ForegroundColor Green
        } catch {
            Write-Host "❌ Erro ao personalizar configuração: $_" -ForegroundColor Red
        }
    }
}

# Função para finalizar instalação
function Complete-Installation {
    Write-Host ""
    Write-Host "Finalizando instalação..." -ForegroundColor Cyan
    
    # Verificar se o script principal existe
    $startScript = Join-Path -Path (Get-Location) -ChildPath "Start-ADRT.ps1"
    if (Test-Path -Path $startScript) {
        # Criar atalho na área de trabalho (opcional)
        $createShortcut = Read-Host "Deseja criar um atalho na área de trabalho? (S/N)"
        if ($createShortcut -eq "S" -or $createShortcut -eq "s") {
            try {
                $desktopPath = [Environment]::GetFolderPath("Desktop")
                $shortcutPath = Join-Path -Path $desktopPath -ChildPath "ADRT Moderno.lnk"
                
                $WshShell = New-Object -ComObject WScript.Shell
                $Shortcut = $WshShell.CreateShortcut($shortcutPath)
                $Shortcut.TargetPath = "powershell.exe"
                $Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$startScript`""
                $Shortcut.WorkingDirectory = (Get-Location).Path
                $Shortcut.IconLocation = "powershell.exe,0"
                $Shortcut.Description = "ADRT - Active Directory Report Tool"
                $Shortcut.Save()
                
                Write-Host "✅ Atalho criado na área de trabalho" -ForegroundColor Green
            } catch {
                Write-Host "❌ Erro ao criar atalho: $_" -ForegroundColor Red
            }
        }
        
        # Perguntar se deseja abrir o ADRT agora
        $openNow = Read-Host "Deseja iniciar o ADRT agora? (S/N)"
        if ($openNow -eq "S" -or $openNow -eq "s") {
            try {
                Write-Host "Iniciando ADRT..." -ForegroundColor Cyan
                & $startScript
                return
            } catch {
                Write-Host "❌ Erro ao iniciar ADRT: $_" -ForegroundColor Red
            }
        }
    } else {
        # Se não encontrar o script principal, abrir o painel via HTML
        $indexPath = Join-Path -Path (Get-Location) -ChildPath "index-modern.html"
        if (Test-Path -Path $indexPath) {
            $openNow = Read-Host "Deseja abrir o painel de controle agora? (S/N)"
            if ($openNow -eq "S" -or $openNow -eq "s") {
                try {
                    Start-Process $indexPath
                } catch {
                    Write-Host "❌ Erro ao abrir o painel de controle: $_" -ForegroundColor Red
                }
            }
        }
    }
    
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                    INSTALAÇÃO CONCLUÍDA                       ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "O ADRT Moderno foi instalado com sucesso!" -ForegroundColor Cyan
    Write-Host "Para iniciar o ADRT, execute o script 'Start-ADRT.ps1'." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Para gerar relatórios completos, execute:" -ForegroundColor Yellow
    Write-Host "  .\analise-completa.ps1" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Ou para visualizar o painel de controle, abra:" -ForegroundColor Yellow
    Write-Host "  .\index-modern.html" -ForegroundColor Yellow
    Write-Host ""
}

# Função principal
function Main {
    Show-Banner
    
    # Verificar se é administrador
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "⚠️ AVISO: Este script não está sendo executado como administrador." -ForegroundColor Yellow
        Write-Host "Algumas funcionalidades podem não funcionar corretamente." -ForegroundColor Yellow
        
        $continueAnyway = Read-Host "Deseja continuar mesmo assim? (S/N)"
        if ($continueAnyway -ne "S" -and $continueAnyway -ne "s") {
            Write-Host "Instalação cancelada. Por favor, execute como administrador e tente novamente." -ForegroundColor Cyan
            return
        }
    }
    
    # Verificar requisitos
    $requirementsMet = Test-Requirements
    if (-not $requirementsMet) {
        Write-Host "Pré-requisitos não atendidos. Correção necessária antes de continuar." -ForegroundColor Red
        Write-Host "Por favor, resolva os problemas indicados e execute o script novamente." -ForegroundColor Cyan
        return
    }
    
    # Confirmar instalação
    Write-Host ""
    Write-Host "Pronto para instalar o ADRT Moderno" -ForegroundColor Cyan
    Write-Host "Diretório de instalação: $(Get-Location)" -ForegroundColor Cyan
    $confirmInstall = Read-Host "Deseja continuar com a instalação? (S/N)"
    if ($confirmInstall -ne "S" -and $confirmInstall -ne "s") {
        Write-Host "Instalação cancelada pelo usuário." -ForegroundColor Cyan
        return
    }
    
    # Criar estrutura de diretórios
    Create-DirectoryStructure
    
    # Baixar arquivos
    Download-Files
    
    # Configurar arquivos
    Configure-Files
    
    # Finalizar instalação
    Complete-Installation
}

# Iniciar script
Main