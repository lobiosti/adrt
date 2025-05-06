#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
.SYNOPSIS
    Inicializador do ADRT - Active Directory Report Tool (Versão Moderna)
.DESCRIPTION
    Este script é o ponto de entrada principal para o ADRT Moderno.
    Oferece um menu interativo para gerar relatórios do Active Directory.
.NOTES
    Autor: Lobios Segurança • Tecnologia • Inovação
#>

# Definir codificação para garantir acentuação correta
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding'] = 'UTF8'

# Diretório onde o script está localizado
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptPath

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
║      ADRT - Active Directory Report Tool (Versão Moderna)     ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Magenta

    # Exibir informações do sistema
    $domainInfo = Get-ADDomain
    $forestInfo = Get-ADForest
    #$dcInfo = Get-ADDomainController -Discover -Service "PrimaryDC"
    
    Write-Host ""
    Write-Host "INFORMAÇÕES DO DOMÍNIO" -ForegroundColor Cyan
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "Domínio: $($domainInfo.DNSRoot)" -ForegroundColor White
    Write-Host "Floresta: $($forestInfo.Name)" -ForegroundColor White
    Write-Host "Nível Funcional: $($domainInfo.DomainMode)" -ForegroundColor White
    if ($dcInfo) {
        Write-Host "Controlador de Domínio: $($dcInfo.HostName)" -ForegroundColor White
    }
    Write-Host "Data: $(Get-Date -Format 'dd/MM/yyyy HH:mm')" -ForegroundColor White
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

# Função para verificar e criar estrutura de diretórios
function Initialize-Environment {
    Write-Host "Verificando ambiente..." -ForegroundColor Cyan
    
    # Verificar módulos necessários
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "✓ Módulo Active Directory carregado com sucesso" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Módulo Active Directory não encontrado. Por favor, instale as ferramentas RSAT." -ForegroundColor Red
        return $false
    }
    
    try {
        $modulePath = Join-Path -Path $scriptPath -ChildPath "modules\ADRT-Modern.psm1"
        if (Test-Path -Path $modulePath) {
            Import-Module $modulePath -ErrorAction Stop
            Write-Host "✓ Módulo ADRT-Modern carregado com sucesso" -ForegroundColor Green
        }
        else {
            Write-Host "✗ Módulo ADRT-Modern não encontrado em: $modulePath" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "✗ Erro ao carregar o módulo ADRT-Modern: $_" -ForegroundColor Red
        return $false
    }
    
    # Criar diretórios de relatório se não existirem
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
        "ad-reports\ad-analysis",
        "web\img"
    )
    
    foreach ($dir in $directories) {
        $dirPath = Join-Path -Path $scriptPath -ChildPath $dir
        if (-not (Test-Path -Path $dirPath)) {
            try {
                New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
                Write-Host "✓ Diretório criado: $dir" -ForegroundColor Green
            }
            catch {
                Write-Host "✗ Erro ao criar diretório $dir : $_" -ForegroundColor Red
            }
        }
    }
    
    # Verificar se os logos estão presentes
    $logoPath = Join-Path -Path $scriptPath -ChildPath "web\img\lobios-logo.png"
    $logoSmallPath = Join-Path -Path $scriptPath -ChildPath "web\img\lobios-logo-small.png"
    
    if (-not (Test-Path -Path $logoPath)) {
        Write-Host "! Logo principal não encontrado: web\img\lobios-logo.png" -ForegroundColor Yellow
    }
    
    if (-not (Test-Path -Path $logoSmallPath)) {
        Write-Host "! Logo pequeno não encontrado: web\img\lobios-logo-small.png" -ForegroundColor Yellow
    }
    
    return $true
}

# Função para exibir o menu principal
function Show-MainMenu {
    Write-Host "MENU PRINCIPAL" -ForegroundColor Cyan
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "1. Atualizar todos os relatórios"
    Write-Host "2. Gerar relatório específico"
    Write-Host "3. Gerar análise completa (ad-all-modern)"
    Write-Host "4. Gerar análise de segurança"
    Write-Host "5. Abrir painel de controle"
    Write-Host "6. Limpar relatórios antigos"
    Write-Host "7. Sair"
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    
    $choice = Read-Host "Selecione uma opção (1-7)"
    
    switch ($choice) {
        "1" { Generate-AllReports }
        "2" { Show-ReportMenu }
        "3" { Generate-AllInOneReport }
        "4" { Generate-SecurityAnalysis }
        "5" { Open-Dashboard }
        "6" { Clear-OldReports -ScriptDirectory $scriptPath }
        "7" { return $false }
        default { 
            Write-Host "Opção inválida. Por favor, tente novamente." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
    
    return $true
}

# Função para exibir menu de relatórios específicos
function Show-ReportMenu {
    Clear-Host
    Show-Banner
    
    Write-Host "MENU DE RELATÓRIOS" -ForegroundColor Cyan
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "1. Todos os Usuários                 (ad-users-modern.ps1)"
    Write-Host "2. Administradores de Domínio        (ad-admins-modern.ps1)"
    Write-Host "3. Administradores Enterprise        (ad-enterprise-admins-modern.ps1)"
    Write-Host "4. Usuários Desativados             (ad-disabled-modern.ps1)"
    Write-Host "5. Último Login                     (ad-lastlogon-modern.ps1)"
    Write-Host "6. Senha Nunca Expira               (ad-neverexpires-modern.ps1)"
    Write-Host "7. Todos os Grupos                  (ad-groups-modern.ps1)"
    Write-Host "8. Membros de Grupos                (ad-membergroups-modern.ps1)"
    Write-Host "9. Todas as OUs                     (ad-ous-modern.ps1)"
    Write-Host "10. Todos os Computadores           (ad-computers-modern.ps1)"
    Write-Host "11. Todos os Servidores             (ad-servers-modern.ps1)"
    Write-Host "12. Controladores de Domínio        (ad-dcs-modern.ps1)"
    Write-Host "13. Todas as GPOs                   (ad-gpos-modern.ps1)"
    Write-Host "14. Inventário Completo             (ad-inventory-modern.ps1)"
    Write-Host "15. Voltar ao menu principal"
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    
    $choice = Read-Host "Selecione uma opção (1-15)"
    
    # Executar o relatório selecionado
    switch ($choice) {
        "1" { Run-Script "ad-users-modern.ps1" }
        "2" { Run-Script "ad-admins-modern.ps1" }
        "3" { Run-Script "ad-enterprise-admins-modern.ps1" }
        "4" { Run-Script "ad-disabled-modern.ps1" }
        "5" { Run-Script "ad-lastlogon-modern.ps1" }
        "6" { Run-Script "ad-neverexpires-modern.ps1" }
        "7" { Run-Script "ad-groups-modern.ps1" }
        "8" { Run-Script "ad-membergroups-modern.ps1" }
        "9" { Run-Script "ad-ous-modern.ps1" }
        "10" { Run-Script "ad-computers-modern.ps1" }
        "11" { Run-Script "ad-servers-modern.ps1" }
        "12" { Run-Script "ad-dcs-modern.ps1" }
        "13" { Run-Script "ad-gpos-modern.ps1" }
        "14" { Run-Script "ad-inventory-modern.ps1" }
        "15" { return }
        default { 
            Write-Host "Opção inválida. Por favor, tente novamente." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-ReportMenu
        }
    }
}

# Função para executar um script específico
function Run-Script {
    param (
        [string]$ScriptName
    )
    
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath $ScriptName
    
    if (Test-Path -Path $scriptPath) {
        try {
            Clear-Host
            Write-Host "Executando $ScriptName..." -ForegroundColor Cyan
            & $scriptPath
            
            Write-Host ""
            Write-Host "Script executado com sucesso! Pressione qualquer tecla para continuar..." -ForegroundColor Green
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        catch {
            Write-Host "Erro ao executar o script: $_" -ForegroundColor Red
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
    else {
        Write-Host "Script não encontrado: $scriptPath" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Função para gerar todos os relatórios
function Generate-AllReports {
    Clear-Host
    Show-Banner
    
    Write-Host "Gerando todos os relatórios..." -ForegroundColor Cyan
    Write-Host "Este processo pode levar alguns minutos. Por favor, aguarde." -ForegroundColor Yellow
    
    $updateScript = Join-Path -Path $PSScriptRoot -ChildPath "atualizar-relatorios.ps1"
    
    if (Test-Path -Path $updateScript) {
        try {
            & $updateScript
            
            Write-Host ""
            Write-Host "Todos os relatórios foram gerados com sucesso!" -ForegroundColor Green
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        catch {
            Write-Host "Erro ao gerar relatórios: $_" -ForegroundColor Red
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
    else {
        Write-Host "Script de atualização não encontrado: $updateScript" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Função para gerar análise de segurança
function Generate-SecurityAnalysis {
    Clear-Host
    Show-Banner
    
    Write-Host "Gerando análise de segurança..." -ForegroundColor Cyan
    Write-Host "Este processo pode levar alguns minutos. Por favor, aguarde." -ForegroundColor Yellow
    
    $analysisScript = Join-Path -Path $PSScriptRoot -ChildPath "analise-completa.ps1"
    
    if (Test-Path -Path $analysisScript) {
        try {
            & $analysisScript
            
            Write-Host ""
            Write-Host "Análise de segurança gerada com sucesso!" -ForegroundColor Green
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        catch {
            Write-Host "Erro ao gerar análise de segurança: $_" -ForegroundColor Red
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
    else {
        Write-Host "Script de análise não encontrado: $analysisScript" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Função para corrigir problemas de acentuação
function Fix-EncodingIssues {
    Clear-Host
    Show-Banner
    
    Write-Host "Corrigindo problemas de acentuação..." -ForegroundColor Cyan
    Write-Host "Este processo irá corrigir problemas de caracteres acentuados em todos os arquivos." -ForegroundColor Yellow
    
    $fixScript = Join-Path -Path $PSScriptRoot -ChildPath "Fix-Encoding.ps1"
    
    if (Test-Path -Path $fixScript) {
        try {
            & $fixScript
            
            Write-Host ""
            Write-Host "Correção de acentuação concluída!" -ForegroundColor Green
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        catch {
            Write-Host "Erro durante correção de acentuação: $_" -ForegroundColor Red
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
    else {
        Write-Host "Script de correção não encontrado: $fixScript" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Função para abrir o painel de controle
function Open-Dashboard {
    Clear-Host
    Show-Banner
    
    Write-Host "Abrindo painel de controle..." -ForegroundColor Cyan
    
    $dashboardPath = Join-Path -Path $PSScriptRoot -ChildPath "index-modern.html"
    
    if (Test-Path -Path $dashboardPath) {
        try {
            Start-Process $dashboardPath
            Write-Host "Painel de controle aberto com sucesso!" -ForegroundColor Green
        }
        catch {
            Write-Host "Erro ao abrir o painel de controle: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Painel de controle não encontrado: $dashboardPath" -ForegroundColor Red
    }
    
    Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Função principal que executa o fluxo principal
function Main {
    Show-Banner
    
    # Verificar e inicializar o ambiente
    $envReady = Initialize-Environment
    if (-not $envReady) {
        Write-Host "Houve problemas na inicialização do ambiente. Alguns recursos podem não funcionar corretamente." -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar mesmo assim, ou feche a janela para sair..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    
    Write-Host "Ambiente inicializado com sucesso!" -ForegroundColor Green
    Start-Sleep -Seconds 1
    
    # Oferecer para abrir o painel de controle automaticamente
    #$openDashboard = Read-Host "Deseja abrir o painel de controle web agora? (S/N)"
    #if ($openDashboard -eq "S" -or $openDashboard -eq "s") {
    #    Open-Dashboard
    #}
    
    # Loop do menu principal
    $keepRunning = $true
    while ($keepRunning) {
        Clear-Host
        Show-Banner
        $keepRunning = Show-MainMenu
    }
    
    # Mensagem de encerramento
    Clear-Host
    Show-Banner
    Write-Host "ADRT - Active Directory Report Tool encerrado." -ForegroundColor Cyan
    Write-Host "Obrigado por utilizar nossa ferramenta!" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function Generate-AllInOneReport {
    Clear-Host
    Show-Banner
    
    Write-Host "Gerando análise completa em um único relatório..." -ForegroundColor Cyan
    Write-Host "Este processo pode levar alguns minutos. Por favor, aguarde." -ForegroundColor Yellow
    
    $allInOneScript = Join-Path -Path $scriptPath -ChildPath "ad-all-modern.ps1"
    
    if (Test-Path -Path $allInOneScript) {
        try {
            & $allInOneScript
            
            Write-Host ""
            Write-Host "Análise completa (all-in-one) gerada com sucesso!" -ForegroundColor Green
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        catch {
            Write-Host "Erro ao gerar análise completa (all-in-one): $_" -ForegroundColor Red
            Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
    else {
        Write-Host "Script de análise completa não encontrado: $allInOneScript" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function Clear-OldReports {
    Clear-Host
    Show-Banner
    
    Write-Host "Limpeza de Relatórios Antigos" -ForegroundColor Cyan
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray

    # Obtém o diretório do script com fallback seguro
    if ($PSScriptRoot) {
        $scriptDir = $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        Write-Host "✗ Não foi possível determinar o diretório do script." -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    $reportsPath = Join-Path -Path $scriptDir -ChildPath "ad-reports"

    if (-not (Test-Path -Path $reportsPath)) {
        Write-Host "O diretório de relatórios não existe: $reportsPath" -ForegroundColor Yellow
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    Write-Host "Esta operação irá APAGAR TODOS os relatórios gerados anteriormente." -ForegroundColor Red
    Write-Host "O conteúdo da pasta $reportsPath será completamente removido." -ForegroundColor Yellow
    Write-Host ""

    $confirmation = Read-Host "Tem certeza que deseja continuar? (S/N)"

    if ($confirmation.ToUpper() -eq "S") {
        try {
            # Remove a pasta completa e todo seu conteúdo
            Remove-Item -Path $reportsPath -Recurse -Force

            # Recria a estrutura de diretórios vazia
            $directories = @(
                "ad-users",
                "ad-admins",
                "ad-enterprise-admins",
                "ad-disabled",
                "ad-lastlogon",
                "ad-neverexpires",
                "ad-groups",
                "ad-membergroups",
                "ad-ous",
                "ad-computers",
                "ad-servers",
                "ad-dcs",
                "ad-gpos",
                "ad-inventory",
                "ad-analysis",
                "ad-all"
            )

            foreach ($dir in $directories) {
                $dirPath = Join-Path -Path $reportsPath -ChildPath $dir
                New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
            }

            Write-Host ""
            Write-Host "✓ Limpeza concluída com sucesso! Todos os relatórios foram removidos." -ForegroundColor Green
        }
        catch {
            Write-Host "✗ Erro ao limpar os relatórios: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host ""
        Write-Host "Operação cancelada pelo usuário." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Iniciar a execução do script
Main
