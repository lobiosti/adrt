<#
.SYNOPSIS
    Active Directory - Relatório Completo (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para executar todos os relatórios do Active Directory
    e gerar um relatório consolidado utilizando o ADRT-Helper.ps1
.NOTES
    Original: ad-all.ps1
    Convertido para formato moderno e otimizado
#>

# Definir codificação para garantir acentuação correta
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding'] = 'UTF8'

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
║        ADRT - Processando Todos os Relatórios                 ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Magenta

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"
$month = Get-Date -Format "MMM"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-all\ad-all-modern.html"
$outputDir = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-all"
$archiveDir = "$outputDir\ad-$date"

# Criar diretórios se não existirem
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "✓ Diretório de saída criado: $outputDir" -ForegroundColor Green
}

if (-not (Test-Path -Path $archiveDir)) {
    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
    Write-Host "✓ Diretório de arquivamento criado: $archiveDir" -ForegroundColor Green
}

# Obter informações de configuração
if (Test-Path -Path "config\config.txt") {
    try {
        $config = Get-Content -Path "config\config.txt" -Encoding UTF8 -ErrorAction Stop
        $company = $config[7]
        $owner = $config[9]
        $smtpServer = $config[11]
        $port = $config[13]
        $from = $config[15]
        $to = $config[17]
        Write-Host "✓ Arquivo de configuração carregado com sucesso" -ForegroundColor Green
    }
    catch {
        Write-Host "! Erro ao ler arquivo de configuração. Usando valores padrão." -ForegroundColor Yellow
        $company = "Lobios"
        $owner = "Administrador"
        $smtpServer = "smtp.example.com"
        $port = "25"
        $from = "adrt@example.com"
        $to = "admin@example.com"
    }
}
else {
    Write-Host "! Arquivo de configuração não encontrado. Usando valores padrão." -ForegroundColor Yellow
    $company = "Lobios"
    $owner = "Administrador"
    $smtpServer = "smtp.example.com"
    $port = "25"
    $from = "adrt@example.com"
    $to = "admin@example.com"
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

# Executar todos os relatórios individuais
$scripts = @(
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

Write-Host ""
Write-Host "Gerando todos os relatórios individuais..." -ForegroundColor Cyan

$scriptErrors = @()

foreach ($script in $scripts) {
    try {
        Write-Host "Executando $script..." -ForegroundColor Yellow
        & ".\$script" -ErrorAction Stop
        Write-Host "✓ $script executado com sucesso" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Erro ao executar $script : $_" -ForegroundColor Red
        $scriptErrors += $script
    }
}

Write-Host ""
Write-Host "Coletando estatísticas para o relatório consolidado..." -ForegroundColor Cyan

# Coletar estatísticas para o relatório consolidado
$stats = @{
    TotalUsers = 0
    DisabledUsers = 0
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
}

try {
    # Usuários
    $stats.TotalUsers = (Get-ADUser -Filter *).Count
    $stats.DisabledUsers = (Search-ADAccount -AccountDisabled -UsersOnly).Count
    $stats.EnabledUsers = $stats.TotalUsers - $stats.DisabledUsers
    
    # LastLogon
    $days = 90
    $timestamp = (Get-Date).AddDays(-($days))
    $stats.LastLogon90Days = (Get-ADUser -Filter {LastLogonTimeStamp -lt $timestamp -and enabled -eq $true} -Properties LastLogonTimeStamp).Count
    
    # PasswordNeverExpires
    $stats.PasswordNeverExpires = (Get-ADUser -Filter * -Properties PasswordNeverExpires | 
                             Where-Object { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true }).Count
    
    # Grupos
    $stats.TotalGroups = (Get-ADGroup -Filter {name -like "*"}).Count
    
    # OUs
    $stats.TotalOUs = (Get-ADOrganizationalUnit -Filter {name -like "*"}).Count
    
    # Computadores
    $stats.TotalComputers = (Get-ADComputer -Filter { OperatingSystem -NotLike '*Windows Server*' }).Count
    
    # Servidores
    $stats.TotalServers = (Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' }).Count
    
    # Total Dispositivos
    $stats.TotalDevices = (Get-ADComputer -Filter *).Count
    
    # Domain Controllers
    $stats.DomainControllers = (Get-ADDomainController -Filter *).Count
    
    # GPOs
    $stats.TotalGPOs = (Get-GPO -All).Count
    
    # Administradores
    $stats.DomainAdmins = (Get-ADGroupMember -Identity "Domain Admins" -ErrorAction SilentlyContinue).Count
    $stats.EnterpriseAdmins = (Get-ADGroupMember -Identity "Enterprise Admins" -ErrorAction SilentlyContinue).Count
    
    # Nome do domínio
    $stats.DomainName = (Get-ADDomain).Forest
    
    Write-Host "✓ Estatísticas coletadas com sucesso:" -ForegroundColor Green
    Write-Host "  - Total de usuários: $($stats.TotalUsers)" -ForegroundColor Gray
    Write-Host "  - Usuários ativos: $($stats.EnabledUsers)" -ForegroundColor Gray
    Write-Host "  - Usuários desativados: $($stats.DisabledUsers)" -ForegroundColor Gray
    Write-Host "  - Usuários com senha que nunca expira: $($stats.PasswordNeverExpires)" -ForegroundColor Gray
    Write-Host "  - Usuários sem login há $days dias: $($stats.LastLogon90Days)" -ForegroundColor Gray
    Write-Host "  - Total de grupos: $($stats.TotalGroups)" -ForegroundColor Gray
    Write-Host "  - Total de OUs: $($stats.TotalOUs)" -ForegroundColor Gray
    Write-Host "  - Total de computadores: $($stats.TotalComputers)" -ForegroundColor Gray
    Write-Host "  - Total de servidores: $($stats.TotalServers)" -ForegroundColor Gray
    Write-Host "  - Total de dispositivos: $($stats.TotalDevices)" -ForegroundColor Gray
    Write-Host "  - Controladores de domínio: $($stats.DomainControllers)" -ForegroundColor Gray
    Write-Host "  - Total de GPOs: $($stats.TotalGPOs)" -ForegroundColor Gray
    Write-Host "  - Administradores de domínio: $($stats.DomainAdmins)" -ForegroundColor Gray
    Write-Host "  - Administradores enterprise: $($stats.EnterpriseAdmins)" -ForegroundColor Gray
}
catch {
    Write-Host "✗ Erro ao coletar estatísticas: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Gerando relatório consolidado..." -ForegroundColor Cyan

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Análise Completa</h1>
    <div class="header-actions">
        <button onclick="window.print()"><i class="fas fa-print"></i> Imprimir</button>
        <button onclick="exportToPDF()"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
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
                <p><strong>Data:</strong> $date</p>
                <p><strong>Responsável:</strong> $owner</p>
            </div>
        </div>
    </div>

    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Status da Geração de Relatórios</div>
            <div class="card-body">
"@

if ($scriptErrors.Count -eq 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Status:</strong> Todos os relatórios foram gerados com sucesso.</p>
                    <p><strong>Total de relatórios gerados:</strong> $($scripts.Count)</p>
                </div>
"@
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>Alguns relatórios não puderam ser gerados ($($scriptErrors.Count))</span>
                </div>
                <div class="info-box">
                    <p><strong>Scripts com erro:</strong></p>
                    <ul>
"@
    
    foreach ($errorScript in $scriptErrors) {
        $bodyContent += @"
                        <li>$errorScript</li>
"@
    }
    
    $bodyContent += @"
                    </ul>
                </div>
"@
}

$bodyContent += @"
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
            <h3>$($stats.TotalDevices)</h3>
            <p>Total de Dispositivos</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-users-cog"></i>
            <h3>$($stats.TotalGroups)</h3>
            <p>Total de Grupos</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-sitemap"></i>
            <h3>$($stats.TotalOUs)</h3>
            <p>Total de OUs</p>
        </div>
    </div>
</div>

<!-- Gráfico de Usuários -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Status de Usuários</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="userStatusChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Status de Dispositivos</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="deviceStatusChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Gráfico de Segurança -->
<div class="card mb-4">
    <div class="card-header">Indicadores de Segurança</div>
    <div class="card-body">
        <div class="row">
            <div class="col-md-8">
                <div class="chart-container">
                    <canvas id="securityChart"></canvas>
                </div>
            </div>
            <div class="col-md-4">
                <h5>Possíveis problemas de segurança:</h5>
                <ul>
"@

if ($stats.PasswordNeverExpires -gt 0) {
    $bodyContent += @"
                    <li>$($stats.PasswordNeverExpires) usuários com senhas que nunca expiram</li>
"@
}

if ($stats.LastLogon90Days -gt 0) {
    $bodyContent += @"
                    <li>$($stats.LastLogon90Days) usuários sem login nos últimos 90 dias</li>
"@
}

if ($stats.DomainAdmins -gt 5) {
    $bodyContent += @"
                    <li>Número elevado de administradores de domínio ($($stats.DomainAdmins))</li>
"@
}

if ($stats.EnterpriseAdmins -gt 2) {
    $bodyContent += @"
                    <li>Número elevado de administradores enterprise ($($stats.EnterpriseAdmins))</li>
"@
}

if ($stats.PasswordNeverExpires -eq 0 -and $stats.LastLogon90Days -eq 0 -and $stats.DomainAdmins -le 5 -and $stats.EnterpriseAdmins -le 2) {
    $bodyContent += @"
                    <li>Nenhum problema de segurança crítico identificado</li>
"@
}

$bodyContent += @"
                </ul>
            </div>
        </div>
    </div>
</div>

<!-- Tabela de Estatísticas -->
<div class="card mb-4">
    <div class="card-header">Estatísticas Completas do Active Directory</div>
    <div class="card-body">
        <div class="table-responsive">
            <table class="table table-bordered table-hover">
                <thead class="table-light">
                    <tr>
                        <th>Descrição</th>
                        <th>Total</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Usuários</td>
                        <td>$($stats.TotalUsers)</td>
                    </tr>
                    <tr>
                        <td>Usuários Ativos</td>
                        <td>$($stats.EnabledUsers)</td>
                    </tr>
                    <tr>
                        <td>Usuários Desativados</td>
                        <td>$($stats.DisabledUsers)</td>
                    </tr>
                    <tr>
                        <td>Usuários sem Login (90 dias)</td>
                        <td>$($stats.LastLogon90Days)</td>
                    </tr>
                    <tr>
                        <td>Senhas que Nunca Expiram</td>
                        <td>$($stats.PasswordNeverExpires)</td>
                    </tr>
                    <tr>
                        <td>Grupos</td>
                        <td>$($stats.TotalGroups)</td>
                    </tr>
                    <tr>
                        <td>Unidades Organizacionais</td>
                        <td>$($stats.TotalOUs)</td>
                    </tr>
                    <tr>
                        <td>Computadores</td>
                        <td>$($stats.TotalComputers)</td>
                    </tr>
                    <tr>
                        <td>Servidores</td>
                        <td>$($stats.TotalServers)</td>
                    </tr>
                    <tr>
                        <td>Total de Dispositivos</td>
                        <td>$($stats.TotalDevices)</td>
                    </tr>
                    <tr>
                        <td>Controladores de Domínio</td>
                        <td>$($stats.DomainControllers)</td>
                    </tr>
                    <tr>
                        <td>GPOs</td>
                        <td>$($stats.TotalGPOs)</td>
                    </tr>
                    <tr>
                        <td>Administradores de Domínio</td>
                        <td>$($stats.DomainAdmins)</td>
                    </tr>
                    <tr>
                        <td>Administradores Enterprise</td>
                        <td>$($stats.EnterpriseAdmins)</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- Links para relatórios individuais -->
<div class="card mb-4">
    <div class="card-header">Relatórios Individuais</div>
    <div class="card-body">
        <div class="row">
            <div class="col-md-6">
                <div class="report-links">
                    <h5>Relatórios de Usuários</h5>
                    <a href="../ad-users/ad-users-modern.html"><i class="fas fa-users"></i> Todos os Usuários</a>
                    <a href="../ad-admins/ad-admins-modern.html"><i class="fas fa-user-shield"></i> Administradores de Domínio</a>
                    <a href="../ad-enterprise-admins/ad-enterprise-admins-modern.html"><i class="fas fa-user-tie"></i> Administradores Enterprise</a>
                    <a href="../ad-disabled/ad-disabled-modern.html"><i class="fas fa-user-times"></i> Usuários Desativados</a>
                    <a href="../ad-lastlogon/ad-lastlogon-modern.html"><i class="fas fa-clock"></i> Último Login</a>
                    <a href="../ad-neverexpires/ad-neverexpires-modern.html"><i class="fas fa-key"></i> Senhas Nunca Expiram</a>
                    <a href="../ad-groups/ad-groups-modern.html"><i class="fas fa-users-cog"></i> Todos os Grupos</a>
                </div>
            </div>
            <div class="col-md-6">
                <div class="report-links">
                    <h5>Relatórios de Infraestrutura</h5>
                    <a href="../ad-membergroups/ad-membergroups-modern.html"><i class="fas fa-layer-group"></i> Membros de Grupos</a>
                    <a href="../ad-ous/ad-ous-modern.html"><i class="fas fa-sitemap"></i> Todas as OUs</a>
                    <a href="../ad-computers/ad-computers-modern.html"><i class="fas fa-desktop"></i> Todos os Computadores</a>
                    <a href="../ad-servers/ad-servers-modern.html"><i class="fas fa-server"></i> Todos os Servidores</a>
                    <a href="../ad-dcs/ad-dcs-modern.html"><i class="fas fa-shield-alt"></i> Controladores de Domínio</a>
                    <a href="../ad-gpos/ad-gpos-modern.html"><i class="fas fa-cogs"></i> Todas as GPOs</a>
                    <a href="../ad-inventory/ad-inventory-modern.html"><i class="fas fa-clipboard-list"></i> Inventário</a>
                </div>
            </div>
        </div>
    </div>
</div>
"@

# Script específico para esta página
$extraScripts = @"
<script>
    // Inicializar os gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de status de usuários
        const userCtx = document.getElementById('userStatusChart').getContext('2d');
        const userChart = new Chart(userCtx, {
            type: 'pie',
            data: {
                labels: ['Usuários Ativos', 'Usuários Desativados', 'Sem Login (90 dias)', 'Senhas Nunca Expiram'],
                datasets: [{
                    data: [$($stats.EnabledUsers - $stats.LastLogon90Days - $stats.PasswordNeverExpires), 
                           $($stats.DisabledUsers), 
                           $($stats.LastLogon90Days), 
                           $($stats.PasswordNeverExpires)],
                    backgroundColor: [
                        '#28a745', // Verde para ativos
                        '#dc3545', // Vermelho para desativados
                        '#ffc107', // Amarelo para sem login
                        '#fd7e14'  // Laranja para senhas não expiram
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
        
        // Gráfico de status de dispositivos
        const deviceCtx = document.getElementById('deviceStatusChart').getContext('2d');
        const deviceChart = new Chart(deviceCtx, {
            type: 'pie',
            data: {
                labels: ['Computadores', 'Servidores', 'Controladores de Domínio'],
                datasets: [{
                    data: [$($stats.TotalComputers), 
                           $($stats.TotalServers - $stats.DomainControllers), 
                           $($stats.DomainControllers)],
                    backgroundColor: [
                        '#9657c7', // Roxo claro para computadores
                        '#6a3094', // Roxo para servidores
                        '#4c1d6b'  // Roxo escuro para DCs
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
        
        // Gráfico de segurança
        const securityCtx = document.getElementById('securityChart').getContext('2d');
        const securityChart = new Chart(securityCtx, {
            type: 'bar',
            data: {
                labels: ['Senhas Nunca Expiram', 'Sem Login (90 dias)', 'Domain Admins', 'Enterprise Admins'],
                datasets: [{
                    label: 'Quantidade',
                    data: [$($stats.PasswordNeverExpires), 
                           $($stats.LastLogon90Days), 
                           $($stats.DomainAdmins), 
                           $($stats.EnterpriseAdmins)],
                    backgroundColor: [
                        '#fd7e14', // Laranja para senhas não expiram
                        '#ffc107', // Amarelo para sem login
                        '#6a3094', // Roxo para Domain Admins
                        '#9657c7'  // Roxo claro para Enterprise Admins
                    ],
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
                }
            }
        });
    });

    // Funções para interação
    function exportToPDF() {
        alert('A função de exportação para PDF está em desenvolvimento.');
        // Implementação da exportação para PDF
    }
    
    function sendReport() {
        alert('A função de envio de relatório por email está em desenvolvimento.');
        // Implementação do envio de relatório
    }
</script>
"@

# Adicionar estilos adicionais para os links de relatórios
$extraScripts += @"
<style>
    .report-links a {
        display: block;
        padding: 10px 15px;
        margin-bottom: 5px;
        background-color: var(--lobios-light);
        color: var(--lobios-primary);
        border-radius: 5px;
        text-decoration: none;
        transition: all 0.3s;
    }
    
    .report-links a:hover {
        background-color: var(--lobios-primary);
        color: white;
    }
    
    .report-links i {
        margin-right: 10px;
        width: 20px;
        text-align: center;
    }
</style>
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
    [System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Host "✓ Relatório HTML consolidado gerado com sucesso em: $outputPath" -ForegroundColor Green
}
catch {
    Write-Host "✗ Erro ao salvar o relatório consolidado: $_" -ForegroundColor Red
}

# Copiar relatórios individuais para o diretório de arquivamento
Write-Host ""
Write-Host "Arquivando relatórios individuais..." -ForegroundColor Cyan

$reportFiles = @(
    "ad-reports\ad-users\ad-users-modern.html",
    "ad-reports\ad-admins\ad-admins-modern.html",
    "ad-reports\ad-enterprise-admins\ad-enterprise-admins-modern.html",
    "ad-reports\ad-disabled\ad-disabled-modern.html",
    "ad-reports\ad-lastlogon\ad-lastlogon-modern.html",
    "ad-reports\ad-neverexpires\ad-neverexpires-modern.html",
    "ad-reports\ad-groups\ad-groups-modern.html",
    "ad-reports\ad-membergroups\ad-membergroups-modern.html",
    "ad-reports\ad-ous\ad-ous-modern.html",
    "ad-reports\ad-computers\ad-computers-modern.html",
    "ad-reports\ad-servers\ad-servers-modern.html",
    "ad-reports\ad-dcs\ad-dcs-modern.html",
    "ad-reports\ad-gpos\ad-gpos-modern.html",
    "ad-reports\ad-gpos\gpos-html.zip",
"ad-reports\ad-inventory\ad-inventory-modern.html"
)

foreach ($file in $reportFiles) {
    if (Test-Path -Path $file) {
        try {
            Copy-Item -Path $file -Destination $archiveDir -Force
            Write-Host "  Copiado: $file" -ForegroundColor Gray
        }
        catch {
            Write-Host "  Erro ao copiar $file : $_" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  Arquivo não encontrado: $file" -ForegroundColor Yellow
    }
}

# Enviando relatório por email (descomentado para uso em produção)

$Subject = "[ Relatório-$month ] Active Directory - Análise Completa"

try {
    $bodyHTML = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Calibri; font-size: 14px; }
        table { margin: auto; width: 80%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        .header { background-color: #6a3094; color: white; padding: 10px; text-align: center; }
        .footer { background-color: #f2f2f2; color: #666; padding: 5px; text-align: center; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Relatório do Active Directory - $date</h2>
    </div>
    
    <p>Prezado(a) $owner,</p>
    
    <p>Segue o relatório completo do Active Directory para o domínio <strong>$($stats.DomainName)</strong>.</p>
    
    <table>
        <tr>
            <th colspan="2">Resumo do Active Directory</th>
        </tr>
        <tr>
            <td>Usuários</td>
            <td>$($stats.TotalUsers)</td>
        </tr>
        <tr>
            <td>Computadores</td>
            <td>$($stats.TotalComputers)</td>
        </tr>
        <tr>
            <td>Servidores</td>
            <td>$($stats.TotalServers)</td>
        </tr>
        <tr>
            <td>Grupos</td>
            <td>$($stats.TotalGroups)</td>
        </tr>
        <tr>
            <td>OUs</td>
            <td>$($stats.TotalOUs)</td>
        </tr>
        <tr>
            <td>GPOs</td>
            <td>$($stats.TotalGPOs)</td>
        </tr>
    </table>
    
    <p>Para acessar o relatório completo, abra o arquivo HTML anexo ou acesse diretamente no servidor.</p>
    
    <p>Atenciosamente,<br/>
    ADRT - Active Directory Report Tool</p>
    
    <div class="footer">
        <p>ADRT - Active Directory Report Tool v2.0 | Desenvolvido por Lobios Segurança • Tecnologia • Inovação</p>
    </div>
</body>
</html>
"@

    $attachments = @(
        "$directoryPath\$outputPath"
    )

    # Adicionar outros relatórios importantes se necessário
    if (Test-Path "$directoryPath\ad-reports\ad-admins\ad-admins-modern.html") {
        $attachments += "$directoryPath\ad-reports\ad-admins\ad-admins-modern.html"
    }
    
    if (Test-Path "$directoryPath\ad-reports\ad-disabled\ad-disabled-modern.html") {
        $attachments += "$directoryPath\ad-reports\ad-disabled\ad-disabled-modern.html"
    }

    Send-MailMessage -From $from -To $to -Subject $Subject -BodyAsHtml -Body $bodyHTML -SmtpServer $smtpServer -Port $port -Attachments $attachments
    Write-Host "✓ Email com relatório enviado com sucesso para $to" -ForegroundColor Green
}
catch {
    Write-Host "✗ Erro ao enviar email: $_" -ForegroundColor Red
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
Write-Host "║                   PROCESSAMENTO CONCLUÍDO                     ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Todos os relatórios foram processados e o relatório consolidado foi gerado."
Write-Host "Os arquivos individuais foram copiados para: $archiveDir"
Write-Host ""
Write-Host "Relatório consolidado: $outputPath"
Write-Host ""