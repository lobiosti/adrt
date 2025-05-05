<#
.SYNOPSIS
    Active Directory - Todos os Servidores (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar servidores no Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-servers.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-servers\ad-servers-modern.html"

# Criar diretório se não existir
$outputDir = Split-Path -Path $outputPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Obter informações de configuração
if (Test-Path -Path "config\config.txt") {
    $config = Get-Content -Path "config\config.txt"
    $company = $config[7]
    $owner = $config[9]
}
else {
    $company = "Lobios"
    $owner = "Administrador"
}

# Carregar o helper
. ".\modules\ADRT-Helper.ps1"

# Importar módulo ActiveDirectory
Import-Module ActiveDirectory

# Coletar dados
Write-Host "Coletando dados de servidores do Active Directory..."
try {
    # Obter todos os servidores Windows
    $servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties Name, OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, Description, IPv4Address, Created, Modified, LastLogonDate, Enabled, DNSHostName, DistinguishedName
    
    # Coletar estatísticas
    $totalServers = $servers.Count
    $enabledServers = ($servers | Where-Object { $_.Enabled -eq $true }).Count
    $disabledServers = $totalServers - $enabledServers

    # Estatísticas por versão do sistema operacional
    $osSummary = @{}
    $servers | ForEach-Object {
        $os = if ($_.OperatingSystem) { 
            # Simplificar nomes de SO
            if ($_.OperatingSystem -like "*Windows Server 2016*") { "Windows Server 2016" }
            elseif ($_.OperatingSystem -like "*Windows Server 2019*") { "Windows Server 2019" }
            elseif ($_.OperatingSystem -like "*Windows Server 2022*") { "Windows Server 2022" }
            else { $_.OperatingSystem }
        } else { "Desconhecido" }
        
        if ($osSummary.ContainsKey($os)) {
            $osSummary[$os]++
        } else {
            $osSummary[$os] = 1
        }
    }
    
    # Verificar servidores que não fizeram login recentemente (90 dias)
    $daysThreshold = 90
    $cutoffDate = (Get-Date).AddDays(-$daysThreshold)
    $inactiveServers = ($servers | Where-Object { 
        ($_.LastLogonDate -lt $cutoffDate -or $_.LastLogonDate -eq $null) -and $_.Enabled -eq $true 
    }).Count
    
    # Obter controladores de domínio para comparação
    $domainControllers = (Get-ADDomainController -Filter *).Count
}
catch {
    Write-Host "Erro ao coletar informações de servidores: $_" -ForegroundColor Red
    $servers = @()
    $totalServers = 0
    $enabledServers = 0
    $disabledServers = 0
    $osSummary = @{}
    $inactiveServers = 0
    $domainControllers = 0
}

# Preparar os dados para o relatório
$serverData = @()
foreach ($server in $servers) {
    # Calcular dias desde o último login
    $daysSinceLogin = if ($server.LastLogonDate) {
        [math]::Round(((Get-Date) - $server.LastLogonDate).TotalDays, 0)
    } else {
        "N/A"
    }
    
    # Determinar status de atividade
    $activityStatus = if (-not $server.Enabled) {
        "Desativado"
    } elseif ($daysSinceLogin -eq "N/A" -or $daysSinceLogin -gt $daysThreshold) {
        "Inativo"
    } else {
        "Ativo"
    }
    
    # Formatar versão do sistema
    $osVersion = if ($server.OperatingSystemVersion) {
        "$($server.OperatingSystem) (Build $($server.OperatingSystemVersion))"
    } else {
        $server.OperatingSystem
    }
    
    # Adicionar ServicePack se disponível
    if ($server.OperatingSystemServicePack) {
        $osVersion += " $($server.OperatingSystemServicePack)"
    }
    
    $serverData += [PSCustomObject]@{
        Nome = $server.Name
        DNS = $server.DNSHostName
        IP = $server.IPv4Address
        SistemaOperacional = $osVersion
        Descricao = $server.Description
        UltimoLogin = if ($server.LastLogonDate) { $server.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        DiasSemLogin = $daysSinceLogin
        Status = $activityStatus
        DataCriacao = if ($server.Created) { $server.Created.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        DN = $server.DistinguishedName
    }
}

# Ordenar resultados
$serverData = $serverData | Sort-Object -Property Nome

# Contar registros
$totalRecords = $serverData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Servidores Windows</h1>
    <div class="header-actions">
        <button onclick="exportToCsv()"><i class="fas fa-file-export"></i> Exportar CSV</button>
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
                <p><strong>Domínio:</strong> $((Get-ADDomain).Forest)</p>
                <p><strong>Data:</strong> $date</p>
                <p><strong>Responsável:</strong> $owner</p>
            </div>
        </div>
    </div>

    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Resumo da Infraestrutura</div>
            <div class="card-body">
"@

if ($totalServers -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Servidores:</strong> $totalServers</p>
                    <p><strong>Servidores Ativos:</strong> $enabledServers</p>
                    <p><strong>Servidores Desativados:</strong> $disabledServers</p>
                    <p><strong>Servidores Inativos (sem login há $daysThreshold dias):</strong> $inactiveServers</p>
                </div>
"@

    if ($inactiveServers -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$inactiveServers servidores sem login nos últimos $daysThreshold dias.</span>
                </div>
"@
    }
    
    if ($totalServers -eq $domainControllers) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Todos os servidores são controladores de domínio. Considere adicionar servidores de aplicação separados.</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum servidor Windows encontrado ou você não tem permissão para visualizar.</span>
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
            <i class="fas fa-server"></i>
            <h3>$totalServers</h3>
            <p>Total de Servidores</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-shield-alt"></i>
            <h3>$domainControllers</h3>
            <p>Controladores de Domínio</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-check-circle text-success"></i>
            <h3>$enabledServers</h3>
            <p>Servidores Ativos</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-power-off text-danger"></i>
            <h3>$disabledServers</h3>
            <p>Servidores Desativados</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribuição por Sistema Operacional</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="osChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Status dos Servidores</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="statusChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Lista de Servidores Windows</div>
        <div>
            <span class="badge badge-primary" id="recordCount">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
"@

if ($totalServers -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="serversTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>DNS/FQDN</th>
                        <th>Endereço IP</th>
                        <th>Sistema Operacional</th>
                        <th>Descrição</th>
                        <th>Último Login</th>
                        <th>Status</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($server in $serverData) {
        $statusBadge = ""
        if ($server.Status -eq "Ativo") {
            $statusBadge = '<span class="badge-status badge-success">Ativo</span>'
        } elseif ($server.Status -eq "Inativo") {
            $statusBadge = '<span class="badge-status badge-warning">Inativo</span>'
        } else {
            $statusBadge = '<span class="badge-status badge-danger">Desativado</span>'
        }
        
        $bodyContent += @"
                    <tr>
                        <td><i class="fas fa-server"></i> $($server.Nome)</td>
                        <td>$($server.DNS)</td>
                        <td>$($server.IP)</td>
                        <td>$($server.SistemaOperacional)</td>
                        <td>$($server.Descricao)</td>
                        <td>$($server.UltimoLogin)</td>
                        <td>$statusBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewServer('$($server.Nome)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="pingServer('$($server.IP)')"><i class="fas fa-network-wired"></i></button>
                            <button class="action-button" onclick="rdpServer('$($server.Nome)')"><i class="fas fa-desktop"></i></button>
                        </td>
                    </tr>
"@
    }

    $bodyContent += @"
                </tbody>
            </table>
        </div>
"@
} else {
    $bodyContent += @"
        <div class="alert alert-info">
            <i class="fas fa-info-circle"></i> Nenhum servidor Windows encontrado neste domínio ou você não tem permissão para visualizar.
        </div>
"@
}

$bodyContent += @"
    </div>
</div>
"@

# Converter os dados do gráfico para JavaScript
$osLabels = "["
$osData = "["
$osColors = "["

$colorPalette = @(
    '#6a3094', '#9657c7', '#8244b2', '#2c1445', '#c9a6e9', 
    '#8e60b0', '#5d2079', '#7f38aa', '#a77bc9', '#543264'
)

$i = 0
foreach ($os in $osSummary.Keys) {
    $osLabels += "'$os',"
    $osData += "$($osSummary[$os]),"
    $osColors += "'$($colorPalette[$i % $colorPalette.Count])',"
    $i++
}

$osLabels = $osLabels.TrimEnd(',') + "]"
$osData = $osData.TrimEnd(',') + "]"
$osColors = $osColors.TrimEnd(',') + "]"

# Script específico para esta página
$extraScripts = @"
<script>
    // Inicializar gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de sistemas operacionais
        const osCtx = document.getElementById('osChart').getContext('2d');
        const osChart = new Chart(osCtx, {
            type: 'pie',
            data: {
                labels: $osLabels,
                datasets: [{
                    data: $osData,
                    backgroundColor: $osColors,
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
        
        // Gráfico de status dos servidores
        const statusCtx = document.getElementById('statusChart').getContext('2d');
        const statusChart = new Chart(statusCtx, {
            type: 'bar',
            data: {
                labels: ['Ativos', 'Inativos', 'Desativados'],
                datasets: [{
                    label: 'Quantidade',
                    data: [
                        $($enabledServers - $inactiveServers),
                        $inactiveServers,
                        $disabledServers
                    ],
                    backgroundColor: [
                        '#28a745', // Verde para ativos
                        '#ffc107', // Amarelo para inativos
                        '#dc3545'  // Vermelho para desativados
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
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#serversTable tbody tr');
            let visibleCount = 0;
            
            rows.forEach(row => {
                let shouldShow = false;
                // Verificar todas as células da linha exceto a última (botões de ação)
                const cells = row.querySelectorAll('td:not(:last-child)');
                cells.forEach(cell => {
                    if (cell.textContent.toLowerCase().includes(searchTerm)) {
                        shouldShow = true;
                    }
                });
                
                // Aplicar visibilidade
                if (shouldShow) {
                    row.style.display = '';
                    visibleCount++;
                } else {
                    row.style.display = 'none';
                }
            });
            
            // Atualizar contagem de registros visíveis
            document.getElementById('recordCount').textContent = visibleCount + ' registros';
        });
    });

    // Funções para interação com servidores
    function viewServer(serverName) {
        alert('Visualizando detalhes do servidor: ' + serverName);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function pingServer(ipAddress) {
        if (!ipAddress || ipAddress === "") {
            alert('Endereço IP não disponível para este servidor');
            return;
        }
        alert('Testando conectividade com: ' + ipAddress);
        // Aqui poderia implementar uma verificação de conectividade real
    }
    
    function rdpServer(serverName) {
        alert('Iniciando conexão RDP para: ' + serverName);
        // Aqui poderia iniciar uma conexão RDP
    }
    
    // Função para exportar para CSV
    function exportToCsv() {
        alert('Exportando dados para CSV...');
        // Implementação da exportação CSV
    }
    
    // Função para enviar relatório
    function sendReport() {
        alert('Enviando relatório por email...');
        // Implementação do envio de relatório
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Servidores" `
                      -ActiveMenu "Servidores" `
                      -CompanyName $company `
                      -DomainName (Get-ADDomain).Forest `
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
    Write-Host "Relatório HTML gerado com sucesso em: $outputPath" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao salvar o relatório: $_" -ForegroundColor Red
}

# Abrir o relatório no navegador
Start-Process $outputPath