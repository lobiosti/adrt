<#
.SYNOPSIS
    Active Directory - Análise Completa (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para análise completa do Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório detalhado
.NOTES
    Original: analise-completa.ps1
    Convertido para formato moderno e otimizado
#>

# Definir codificação para garantir acentuação correta
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding'] = 'UTF8'

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-analysis\ad-analysis-modern.html"

# Criar diretório se não existir
$outputDir = Split-Path -Path $outputPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "✓ Diretório de saída criado: $outputDir" -ForegroundColor Green
}

# Banner
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
║        ADRT - Análise Completa do Active Directory            ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Magenta

# Obter informações de configuração
if (Test-Path -Path "config\config.txt") {
    try {
        $config = Get-Content -Path "config\config.txt" -Encoding UTF8 -ErrorAction Stop
        $company = $config[7]
        $owner = $config[9]
        Write-Host "✓ Arquivo de configuração carregado com sucesso" -ForegroundColor Green
    }
    catch {
        Write-Host "! Erro ao ler arquivo de configuração. Usando valores padrão." -ForegroundColor Yellow
        $company = "Lobios"
        $owner = "Administrador"
    }
}
else {
    Write-Host "! Arquivo de configuração não encontrado. Usando valores padrão." -ForegroundColor Yellow
    $company = "Lobios"
    $owner = "Administrador"
}

# Carregar o helper
. ".\modules\ADRT-Helper.ps1"

# Importar módulo ActiveDirectory
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "✓ Módulo ActiveDirectory carregado com sucesso" -ForegroundColor Green
}
catch {
    Write-Host "✗ Erro crítico: Não foi possível carregar o módulo ActiveDirectory" -ForegroundColor Red
    Write-Host "Este script requer o módulo ActiveDirectory. Verifique se as ferramentas RSAT estão instaladas." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "Iniciando análise completa do Active Directory..." -ForegroundColor Cyan
Write-Host "Coletando estatísticas e métricas..." -ForegroundColor Cyan

# Coletar dados principais - inicializar estrutura
$stats = @{
    TotalUsers = 0
    EnabledUsers = 0
    DisabledUsers = 0
    Days = 90
    LastLogon90Days = 0
    PasswordNeverExpires = 0
    TotalComputers = 0
    TotalServers = 0
    TotalGroups = 0
    TotalOUs = 0
    DomainAdmins = 0
    EnterpriseAdmins = 0
    DomainControllers = 0
    TotalGPOs = 0
    DomainName = ""
    ForestLevel = ""
    DomainLevel = ""
}

# Contagens básicas com tratamento de erro
try {
    $stats.TotalUsers = (Get-ADUser -Filter *).Count
    $stats.EnabledUsers = (Get-ADUser -Filter {Enabled -eq $true}).Count
    $stats.DisabledUsers = $stats.TotalUsers - $stats.EnabledUsers
    Write-Host "Total de usuários: $($stats.TotalUsers)" -ForegroundColor Green
    Write-Host "Usuários ativos: $($stats.EnabledUsers)" -ForegroundColor Green
    Write-Host "Usuários desativados: $($stats.DisabledUsers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar usuários: $_" -ForegroundColor Yellow
}

# Calcular usuários com senha nunca expira
try {
    $stats.PasswordNeverExpires = (Get-ADUser -filter * -properties PasswordNeverExpires | 
        Where-Object { $_.PasswordNeverExpires -eq "true" -and $_.enabled -eq "true" }).Count
    Write-Host "Usuários com senha que nunca expira: $($stats.PasswordNeverExpires)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar usuários com senha que nunca expira: $_" -ForegroundColor Yellow
}

# Calcular usuários sem login nos últimos 90 dias
try {
    $timestamp = (Get-Date).AddDays(-($stats.Days))
    $stats.LastLogon90Days = (Get-ADUser -Filter {LastLogonTimeStamp -lt $timestamp -and enabled -eq $true} -Properties LastLogonTimeStamp).Count
    Write-Host "Usuários sem login nos últimos 90 dias: $($stats.LastLogon90Days)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar usuários sem login recente: $_" -ForegroundColor Yellow
}

# Computadores e servidores
try {
    $stats.TotalComputers = (Get-ADComputer -Filter { OperatingSystem -NotLike '*Windows Server*' }).Count
    Write-Host "Total de computadores: $($stats.TotalComputers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar computadores: $_" -ForegroundColor Yellow
}

try {
    $stats.TotalServers = (Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' }).Count
    Write-Host "Total de servidores: $($stats.TotalServers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar servidores: $_" -ForegroundColor Yellow
}

# Grupos e OUs
try {
    $stats.TotalGroups = (Get-ADGroup -Filter {name -like "*"}).Count
    Write-Host "Total de grupos: $($stats.TotalGroups)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar grupos: $_" -ForegroundColor Yellow
}

try {
    $stats.TotalOUs = (Get-ADOrganizationalUnit -Filter {name -like "*"}).Count
    Write-Host "Total de OUs: $($stats.TotalOUs)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar OUs: $_" -ForegroundColor Yellow
}

# Domain Controllers
try {
    $domainControllers = Get-ADDomainController -Filter * 
    $stats.DomainControllers = $domainControllers.Count
    Write-Host "Total de controladores de domínio: $($stats.DomainControllers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar controladores de domínio: $_" -ForegroundColor Yellow
}

# Obter informações de domínio
try {
    $domain = Get-ADDomain
    $forest = Get-ADForest
    $stats.DomainName = $domain.DNSRoot
    $stats.DomainLevel = $domain.DomainMode
    $stats.ForestLevel = $forest.ForestMode
    Write-Host "Domínio: $($stats.DomainName)" -ForegroundColor Green
    Write-Host "Nível funcional de domínio: $($stats.DomainLevel)" -ForegroundColor Green
    Write-Host "Nível funcional de floresta: $($stats.ForestLevel)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao obter informações de domínio: $_" -ForegroundColor Yellow
}

# Domain Admins
try {
    $stats.DomainAdmins = (Get-ADGroupMember -Identity "Domain Admins" -ErrorAction SilentlyContinue).Count
    Write-Host "Total de administradores de domínio: $($stats.DomainAdmins)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar administradores de domínio: $_" -ForegroundColor Yellow
}

# Enterprise Admins
try {
    $stats.EnterpriseAdmins = (Get-ADGroupMember -Identity "Enterprise Admins" -ErrorAction SilentlyContinue).Count
    Write-Host "Total de administradores enterprise: $($stats.EnterpriseAdmins)" -ForegroundColor Green
}
catch {
    $stats.EnterpriseAdmins = 0
    Write-Host "Grupo Enterprise Admins não encontrado ou erro ao contar" -ForegroundColor Yellow
}

# GPOs
try {
    $stats.TotalGPOs = (Get-GPO -All).Count
    Write-Host "Total de GPOs: $($stats.TotalGPOs)" -ForegroundColor Green
}
catch {
    $stats.TotalGPOs = 0
    Write-Host "Erro ao contar GPOs: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Analisando sistemas operacionais..." -ForegroundColor Cyan

# Análise de sistemas operacionais
$osList = @{}
try {
    $computers = Get-ADComputer -Filter * -Properties OperatingSystem
    
    foreach ($computer in $computers) {
        if ($computer.OperatingSystem) {
            $os = $computer.OperatingSystem
            # Simplificar nomes para agrupamento
            if ($os -like "*Windows 10*") { $os = "Windows 10" }
            elseif ($os -like "*Windows 11*") { $os = "Windows 11" }
            elseif ($os -like "*Windows Server 2016*") { $os = "Windows Server 2016" }
            elseif ($os -like "*Windows Server 2019*") { $os = "Windows Server 2019" }
            elseif ($os -like "*Windows Server 2022*") { $os = "Windows Server 2022" }
            
            if ($osList.ContainsKey($os)) {
                $osList[$os]++
            } else {
                $osList[$os] = 1
            }
        } else {
            if ($osList.ContainsKey("Desconhecido")) {
                $osList["Desconhecido"]++
            } else {
                $osList["Desconhecido"] = 1
            }
        }
    }
    
    # Exibir resumo de SO
    Write-Host "Distribuição de sistemas operacionais:" -ForegroundColor Green
    foreach ($key in $osList.Keys) {
        Write-Host "  $key : $($osList[$key])" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Erro ao analisar sistemas operacionais: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Realizando análise de segurança..." -ForegroundColor Cyan

# Análise de segurança
$securityAnalysis = @{
    RiskLevel = "Médio"
    Findings = @()
    Recommendations = @()
}

# Verificar porcentagem de contas com senha que nunca expira
if ($stats.TotalUsers -gt 0) {
    $neverExpiresPercentage = [math]::Round(($stats.PasswordNeverExpires / $stats.EnabledUsers) * 100, 1)
    if ($neverExpiresPercentage -gt 10) {
        $securityAnalysis.Findings += "Alto percentual de contas ($neverExpiresPercentage%) com senhas que nunca expiram"
        $securityAnalysis.Recommendations += "Revisar política de senhas e configurar expiração para contas não críticas"
    }
}

# Verificar porcentagem de contas sem login recente
if ($stats.EnabledUsers -gt 0) {
    $inactivePercentage = [math]::Round(($stats.LastLogon90Days / $stats.EnabledUsers) * 100, 1)
    if ($inactivePercentage -gt 15) {
        $securityAnalysis.Findings += "Alto percentual de contas ativas ($inactivePercentage%) sem login nos últimos 90 dias"
        $securityAnalysis.Recommendations += "Revisar e desativar contas inativas para reduzir superfície de ataque"
    }
}

# Verificar número de administradores de domínio
if ($stats.DomainAdmins -gt 5) {
    $securityAnalysis.Findings += "Número elevado de administradores de domínio ($($stats.DomainAdmins))"
    $securityAnalysis.Recommendations += "Reduzir o número de contas com privilégios de administrador de domínio"
}

# Verificar se há Enterprise Admins (se aplicável)
if ($stats.EnterpriseAdmins -gt 2) {
    $securityAnalysis.Findings += "Número elevado de administradores enterprise ($($stats.EnterpriseAdmins))"
    $securityAnalysis.Recommendations += "Restringir os privilégios de administrador enterprise ao mínimo necessário"
}

# Verificar nível funcional de domínio
$outdatedDomainLevel = $false
if ($stats.DomainLevel -like "*2008*" -or $stats.DomainLevel -like "*2003*" -or $stats.DomainLevel -like "*2000*") {
    $outdatedDomainLevel = $true
    $securityAnalysis.Findings += "Nível funcional de domínio desatualizado ($($stats.DomainLevel))"
    $securityAnalysis.Recommendations += "Atualizar o nível funcional do domínio para uma versão mais recente"
}

# Verificar número de controladores de domínio
if ($stats.DomainControllers -lt 2) {
    $securityAnalysis.Findings += "Apenas um controlador de domínio detectado"
    $securityAnalysis.Recommendations += "Implementar pelo menos dois controladores de domínio para redundância"
}

# Definir nível de risco geral com base na quantidade de problemas encontrados
if ($securityAnalysis.Findings.Count -ge 3) {
    $securityAnalysis.RiskLevel = "Alto"
} elseif ($securityAnalysis.Findings.Count -ge 1) {
    $securityAnalysis.RiskLevel = "Médio"
} else {
    $securityAnalysis.RiskLevel = "Baixo"
}

Write-Host "Nível de risco identificado: $($securityAnalysis.RiskLevel)" -ForegroundColor Cyan
foreach ($finding in $securityAnalysis.Findings) {
    Write-Host "  Problema: $finding" -ForegroundColor Yellow
}
foreach ($recommendation in $securityAnalysis.Recommendations) {
    Write-Host "  Recomendação: $recommendation" -ForegroundColor Green
}

Write-Host ""
Write-Host "Coletando dados para relatório..." -ForegroundColor Cyan

# Coletar dados para as tabelas
$topUsers = @()
$disabledUsers = @()
$adminUsers = @()
$servers = @()

try {
    $topUsers = Get-ADUser -Filter * -Properties Name, SamAccountName, EmailAddress, Department, Title, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires | 
        Select-Object Name, SamAccountName, EmailAddress, Department, Title, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires |
        Sort-Object -Property LastLogonDate -Descending |
        Select-Object -First 10
}
catch {
    Write-Host "Erro ao coletar dados de usuários ativos: $_" -ForegroundColor Yellow
}

try {
    $disabledUsers = Search-ADAccount -AccountDisabled -UsersOnly |
        Get-ADUser -Properties Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate |
        Select-Object Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate |
        Sort-Object -Property LastLogonDate -Descending |
        Select-Object -First 10
}
catch {
    Write-Host "Erro ao coletar dados de usuários desativados: $_" -ForegroundColor Yellow
}

try {
    $adminUsers = Get-ADGroupMember -Identity "Domain Admins" |
        Get-ADUser -Properties Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate, PasswordLastSet |
        Select-Object Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate, PasswordLastSet
}
catch {
    Write-Host "Erro ao coletar dados de administradores: $_" -ForegroundColor Yellow
}

try {
    $servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties Name, OperatingSystem, IPv4Address, Description, LastLogonDate |
        Select-Object Name, OperatingSystem, IPv4Address, Description, LastLogonDate |
        Sort-Object -Property LastLogonDate -Descending |
        Select-Object -First 10
}
catch {
    Write-Host "Erro ao coletar dados de servidores: $_" -ForegroundColor Yellow
}

# Preparar dados para o relatório
$userDataHtml = ""
foreach ($user in $topUsers) {
    $statusBadge = ""
    if ($user.Enabled) {
        $statusBadge = '<span class="badge-status badge-success">Ativo</span>'
    } else {
        $statusBadge = '<span class="badge-status badge-danger">Desativado</span>'
    }
    
    $passwordBadge = ""
    if ($user.PasswordNeverExpires) {
        $passwordBadge = '<span class="badge-status badge-warning">Sim</span>'
    } else {
        $passwordBadge = '<span class="badge-status badge-success">Não</span>'
    }
    
    $userDataHtml += @"
                            <tr>
                                <td>$($user.Name)</td>
                                <td>$($user.SamAccountName)</td>
                                <td>$($user.EmailAddress)</td>
                                <td>$($user.Department)</td>
                                <td>$($user.Title)</td>
                                <td>$(if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" })</td>
                                <td>$statusBadge</td>
                                <td>$passwordBadge</td>
                            </tr>
"@
}

$adminDataHtml = ""
foreach ($admin in $adminUsers) {
    $adminDataHtml += @"
                            <tr>
                                <td>$($admin.Name)</td>
                                <td>$($admin.SamAccountName)</td>
                                <td>$($admin.EmailAddress)</td>
                                <td>$($admin.Department)</td>
                                <td>$($admin.Title)</td>
                                <td>$(if ($admin.LastLogonDate) { $admin.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" })</td>
                                <td>$(if ($admin.PasswordLastSet) { $admin.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Desconhecido" })</td>
                            </tr>
"@
}

$serverDataHtml = ""
foreach ($server in $servers) {
    $serverDataHtml += @"
                            <tr>
                                <td>$($server.Name)</td>
                                <td>$($server.OperatingSystem)</td>
                                <td>$($server.IPv4Address)</td>
                                <td>$($server.Description)</td>
                                <td>$(if ($server.LastLogonDate) { $server.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" })</td>
                            </tr>
"@
}

$findingsHtml = ""
foreach ($finding in $securityAnalysis.Findings) {
    $findingsHtml += @"
                        <div class="warning-item warning-yellow">
                            <i class="fas fa-exclamation-triangle"></i>
                            <span>$finding</span>
                        </div>
"@
}

$recommendationsHtml = ""
foreach ($recommendation in $securityAnalysis.Recommendations) {
    $recommendationsHtml += @"
                        <div class="warning-item warning-green">
                            <i class="fas fa-check-circle"></i>
                            <span>$recommendation</span>
                        </div>
"@
}

# Gerar sistema operacional para o gráfico
$osLabelsJs = "["
$osDataJs = "["
$osColorsJs = "["

$colorPalette = @(
    '#6a3094', '#9657c7', '#8244b2', '#2c1445', '#c9a6e9', 
    '#5d2683', '#4c1d6b', '#e0cdf0', '#b688df', '#7e3fa8'
)

$i = 0
foreach ($os in $osList.Keys) {
    $osLabelsJs += "'$os',"
    $osDataJs += "$($osList[$os]),"
    $osColorsJs += "'$($colorPalette[$i % $colorPalette.Count])',"
    $i++
}

$osLabelsJs = $osLabelsJs.TrimEnd(',') + "]"
$osDataJs = $osDataJs.TrimEnd(',') + "]"
$osColorsJs = $osColorsJs.TrimEnd(',') + "]"

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Análise Completa</h1>
    <div class="header-actions">
        <button onclick="exportToPdf()"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
        <button onclick="window.print()"><i class="fas fa-print"></i> Imprimir</button>
        <button onclick="sendReport()"><i class="fas fa-envelope"></i> Enviar Relatório</button>
    </div>
</div>

<!-- Dashboard -->
<div class="row">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Informações do Domínio</div>
            <div class="card-body">
                <p><strong>Empresa:</strong> $company</p>
                <p><strong>Domínio:</strong> $($stats.DomainName)</p>
                <p><strong>Nível Funcional de Domínio:</strong> $($stats.DomainLevel)</p>
                <p><strong>Nível Funcional de Floresta:</strong> $($stats.ForestLevel)</p>
                <p><strong>Data:</strong> $date</p>
                <p><strong>Responsável:</strong> $owner</p>
            </div>
        </div>
    </div>

    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">
                <span>Resumo da Segurança</span>
                <span class="risk-badge risk-badge-$($securityAnalysis.RiskLevel.ToLower())">Risco $($securityAnalysis.RiskLevel)</span>
            </div>
            <div class="card-body">
                <div class="info-box">
                    <p><strong>Total de Usuários:</strong> $($stats.TotalUsers) (Ativos: $($stats.EnabledUsers), Desativados: $($stats.DisabledUsers))</p>
                    <p><strong>Usuários com senha que nunca expira:</strong> $($stats.PasswordNeverExpires)</p>
                    <p><strong>Usuários sem login nos últimos $($stats.Days) dias:</strong> $($stats.LastLogon90Days)</p>
                    <p><strong>Administradores de Domínio:</strong> $($stats.DomainAdmins)</p>
                </div>
                
                <!-- Problemas encontrados -->
$findingsHtml
                
                <!-- Recomendações -->
$recommendationsHtml
            </div>
        </div>
    </div>
</div>

<!-- Estatísticas -->
<div class="row mb-4">
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-users"></i>
            <h3>$($stats.TotalUsers)</h3>
            <p>Total de Usuários</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-desktop"></i>
            <h3>$($stats.TotalComputers)</h3>
            <p>Computadores</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-server"></i>
            <h3>$($stats.TotalServers)</h3>
            <p>Servidores</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-shield-alt"></i>
            <h3>$($stats.DomainControllers)</h3>
            <p>Controladores de Domínio</p>
        </div>
    </div>
</div>

<!-- Gráficos de Resumo -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribuição de Usuários</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="userChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribuição de Sistemas Operacionais</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="osChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabs para navegação de dados -->
<div class="card">
    <div class="card-header">
        <ul class="nav nav-tabs card-header-tabs" id="dataTabs" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active" id="users-tab" data-bs-toggle="tab" data-bs-target="#users" type="button" role="tab" aria-controls="users" aria-selected="true">Usuários Ativos</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="admins-tab" data-bs-toggle="tab" data-bs-target="#admins" type="button" role="tab" aria-controls="admins" aria-selected="false">Administradores</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="servers-tab" data-bs-toggle="tab" data-bs-target="#servers" type="button" role="tab" aria-controls="servers" aria-selected="false">Servidores</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="security-tab" data-bs-toggle="tab" data-bs-target="#security" type="button" role="tab" aria-controls="security" aria-selected="false">Segurança</button>
            </li>
        </ul>
    </div>
    <div class="card-body">
        <div class="tab-content" id="dataTabsContent">
            <!-- Tab de Usuários Ativos -->
            <div class="tab-pane fade show active" id="users" role="tabpanel" aria-labelledby="users-tab">
                <h4>Usuários Ativos Recentes (Top 10)</h4>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Nome</th>
                                <th>Login</th>
                                <th>Email</th>
                                <th>Departamento</th>
                                <th>Cargo</th>
                                <th>Último Login</th>
                                <th>Status</th>
                                <th>Senha Nunca Expira</th>
                            </tr>
                        </thead>
                        <tbody>
$userDataHtml
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Tab de Administradores -->
            <div class="tab-pane fade" id="admins" role="tabpanel" aria-labelledby="admins-tab">
                <h4>Administradores de Domínio</h4>
                <div class="table-responsive">
                    <table>
<thead>
                            <tr>
                                <th>Nome</th>
                                <th>Login</th>
                                <th>Email</th>
                                <th>Departamento</th>
                                <th>Cargo</th>
                                <th>Último Login</th>
                                <th>Última Troca de Senha</th>
                            </tr>
                        </thead>
                        <tbody>
$adminDataHtml
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Tab de Servidores -->
            <div class="tab-pane fade" id="servers" role="tabpanel" aria-labelledby="servers-tab">
                <h4>Servidores Windows</h4>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Nome</th>
                                <th>Sistema Operacional</th>
                                <th>Endereço IP</th>
                                <th>Descrição</th>
                                <th>Último Login</th>
                            </tr>
                        </thead>
                        <tbody>
$serverDataHtml
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Tab de Segurança -->
            <div class="tab-pane fade" id="security" role="tabpanel" aria-labelledby="security-tab">
                <h4>Análise de Segurança</h4>
                
                <div class="mb-4">
                    <h5>Nível de Risco: <span class="risk-$($securityAnalysis.RiskLevel.ToLower())">$($securityAnalysis.RiskLevel)</span></h5>
                    
                    <div class="row">
                        <div class="col-md-6">
                            <div class="card mb-3">
                                <div class="card-header">Problemas Encontrados</div>
                                <div class="card-body">
                                    <ul class="list-group">
$(foreach ($finding in $securityAnalysis.Findings) {
    "                                        <li class='list-group-item'><i class='fas fa-exclamation-triangle text-warning'></i> $finding</li>"
})
                                    </ul>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-6">
                            <div class="card mb-3">
                                <div class="card-header">Recomendações</div>
                                <div class="card-body">
                                    <ul class="list-group">
$(foreach ($recommendation in $securityAnalysis.Recommendations) {
    "                                        <li class='list-group-item'><i class='fas fa-check-circle text-success'></i> $recommendation</li>"
})
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">Métricas de Segurança</div>
                    <div class="card-body">
                        <div class="row">
                            <div class="col-md-4">
                                <div class="stat-card">
                                    <i class="fas fa-key text-warning"></i>
                                    <h3>$($stats.PasswordNeverExpires)</h3>
                                    <p>Senhas Nunca Expiram</p>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="stat-card">
                                    <i class="fas fa-user-clock text-danger"></i>
                                    <h3>$($stats.LastLogon90Days)</h3>
                                    <p>Sem Login (90 dias)</p>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="stat-card">
                                    <i class="fas fa-user-shield text-primary"></i>
                                    <h3>$($stats.DomainAdmins)</h3>
                                    <p>Administradores</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
"@

# Estilos adicionais para as badges de risco
$extraStyles = @"
<style>
    /* Badges de risco */
    .risk-badge {
        display: inline-block;
        padding: 8px 15px;
        border-radius: 30px;
        font-size: 14px;
        font-weight: 600;
        margin-left: 10px;
    }
    
    .risk-badge-baixo {
        background-color: #d4edda;
        color: #155724;
    }
    
    .risk-badge-medio {
        background-color: #fff3cd;
        color: #856404;
    }
    
    .risk-badge-alto {
        background-color: #f8d7da;
        color: #721c24;
    }
    
    /* Cores para texto de risco */
    .risk-baixo {
        color: #28a745;
        font-weight: bold;
    }
    
    .risk-medio {
        color: #ffc107;
        font-weight: bold;
    }
    
    .risk-alto {
        color: #dc3545;
        font-weight: bold;
    }
    
    /* Estilos para as tabs */
    .nav-tabs .nav-link {
        color: var(--lobios-primary);
        border: none;
        padding: 10px 20px;
        border-radius: 0;
        font-weight: 500;
    }
    
    .nav-tabs .nav-link.active {
        color: white;
        background-color: var(--lobios-primary);
        border-bottom: 2px solid var(--lobios-accent);
    }
    
    .nav-tabs .nav-link:hover {
        border-bottom: 2px solid var(--lobios-light);
    }
    
    .tab-content {
        padding: 20px 0;
    }
</style>
"@

# Script específico para esta página
$extraScripts = @"
$extraStyles

<script>
    // Inicializar gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de distribuição de usuários
        const userCtx = document.getElementById('userChart').getContext('2d');
        const userChart = new Chart(userCtx, {
            type: 'pie',
            data: {
                labels: ['Usuários Ativos', 'Usuários Desativados', 'Senhas Nunca Expiram', 'Sem Login (90 dias)'],
                datasets: [{
                    data: [
                        $($stats.EnabledUsers - $stats.PasswordNeverExpires - $stats.LastLogon90Days), 
                        $($stats.DisabledUsers), 
                        $($stats.PasswordNeverExpires), 
                        $($stats.LastLogon90Days)
                    ],
                    backgroundColor: [
                        '#28a745', // Verde para ativos sem problemas
                        '#dc3545', // Vermelho para desativados
                        '#ffc107', // Amarelo para senhas que nunca expiram
                        '#6c757d'  // Cinza para sem login
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right'
                    }
                }
            }
        });
        
        // Gráfico de distribuição de sistemas operacionais
        const osCtx = document.getElementById('osChart').getContext('2d');
        const osChart = new Chart(osCtx, {
            type: 'bar',
            data: {
                labels: $osLabelsJs,
                datasets: [{
                    label: 'Quantidade',
                    data: $osDataJs,
                    backgroundColor: $osColorsJs,
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });
    });

    // Funções para interação
    function exportToPdf() {
        alert('Exportando relatório para PDF...');
        // Implementação da exportação para PDF
    }
    
    function sendReport() {
        alert('Enviando relatório por email...');
        // Implementação do envio de relatório
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Análise Completa" `
                      -ActiveMenu "Análise Completa" `
                      -CompanyName $company `
                      -DomainName $stats.DomainName `
                      -Date $date `
                      -Owner $owner `
                      -ExtraScripts $extraScripts

# Salvar o HTML no arquivo de saída
try {
    # Criar diretório se não existir
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    [System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Host "✓ Relatório HTML gerado com sucesso em: $outputPath" -ForegroundColor Green
}
catch {
    Write-Host "✗ Erro ao salvar o relatório: $_" -ForegroundColor Red
    exit 1
}

# Abrir o relatório no navegador
try {
    Start-Process $outputPath
    Write-Host "✓ Relatório aberto no navegador com sucesso" -ForegroundColor Green
}
catch {
    Write-Host "! Erro ao abrir o relatório no navegador: $_" -ForegroundColor Yellow
    Write-Host "Você pode abrir manualmente o arquivo em: $outputPath" -ForegroundColor Yellow
}

# Fim do script
Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                      ANÁLISE CONCLUÍDA                        ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "A análise completa do Active Directory foi concluída com sucesso."
Write-Host "O relatório foi gerado e aberto no seu navegador padrão."
Write-Host ""
Write-Host "Caminho do relatório: $outputPath"
Write-Host ""
Write-Host "Recomendação: Execute esta análise periodicamente para monitorar"
Write-Host "a segurança e a integridade do seu Active Directory."
Write-Host ""