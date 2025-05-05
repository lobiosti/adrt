<#
.SYNOPSIS
    Active Directory - Inventário Completo (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para inventário completo de computadores e servidores
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-inventory.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-inventory\ad-inventory-modern.html"

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
Write-Host "Coletando dados de inventário de computadores e servidores..."
try {
    # Obter todos os computadores e servidores
    $computers = Get-ADComputer -Filter 'Name -like "*"' -Properties Name, IPv4Address, LastLogonDate, OperatingSystem, OperatingSystemVersion, Description, Enabled, Created, Modified, PasswordLastSet
    
    # Coletar estatísticas para o dashboard
    $totalDevices = $computers.Count
    $totalServers = ($computers | Where-Object { $_.OperatingSystem -like "*Server*" }).Count
    $totalWorkstations = $totalDevices - $totalServers
    
    # Sistemas operacionais
    $osList = @{}
    foreach ($device in $computers) {
        $os = "Desconhecido"
        if ($device.OperatingSystem) {
            $os = $device.OperatingSystem
            # Simplificar nomes de SO
            if ($os -like "*Windows 10*") { $os = "Windows 10" }
            elseif ($os -like "*Windows 11*") { $os = "Windows 11" }
            elseif ($os -like "*Windows Server 2016*") { $os = "Windows Server 2016" }
            elseif ($os -like "*Windows Server 2019*") { $os = "Windows Server 2019" }
            elseif ($os -like "*Windows Server 2022*") { $os = "Windows Server 2022" }
        }
        
        if ($osList.ContainsKey($os)) {
            $osList[$os]++
        } else {
            $osList[$os] = 1
        }
    }
    
    # Informações sobre status
    $disabledDevices = ($computers | Where-Object { $_.Enabled -eq $false }).Count
    
    # Verificar dispositivos que não fizeram login recentemente (90 dias)
    $cutoffDate = (Get-Date).AddDays(-90)
    $inactiveDevices = ($computers | Where-Object { 
        ($_.LastLogonDate -lt $cutoffDate -or $_.LastLogonDate -eq $null) -and $_.Enabled -eq $true 
    }).Count
}
catch {
    Write-Host "Erro ao coletar informações de inventário: $_" -ForegroundColor Red
    $computers = @()
    $totalDevices = 0
    $totalServers = 0
    $totalWorkstations = 0
    $osList = @{}
    $disabledDevices = 0
    $inactiveDevices = 0
}

# Preparar os dados para o relatório
$inventoryData = @()
foreach ($device in $computers) {
    $inventoryData += [PSCustomObject]@{
        Nome = $device.Name
        IP = $device.IPv4Address
        UltimoLogin = if ($device.LastLogonDate) { $device.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        SistemaOperacional = $device.OperatingSystem
        Versao = $device.OperatingSystemVersion
        Descricao = $device.Description
        Status = if ($device.Enabled) { "Ativo" } else { "Inativo" }
        Criado = if ($device.Created) { $device.Created.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        UltimaModificacao = if ($device.Modified) { $device.Modified.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        TipoDispositivo = if ($device.OperatingSystem -like "*Server*") { "Servidor" } else { "Estação de Trabalho" }
    }
}

# Ordenar resultados
$inventoryData = $inventoryData | Sort-Object -Property "SistemaOperacional", "Nome"

# Contar registros
$totalRecords = $inventoryData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Inventário Completo</h1>
    <div class="header-actions">
        <button onclick="filterDevices('all')"><i class="fas fa-sync"></i> Mostrar Todos</button>
        <button onclick="filterDevices('server')"><i class="fas fa-server"></i> Servidores</button>
        <button onclick="filterDevices('workstation')"><i class="fas fa-desktop"></i> Estações</button>
        <button onclick="exportToCsv()"><i class="fas fa-file-export"></i> Exportar CSV</button>
        <button onclick="window.print()"><i class="fas fa-print"></i> Imprimir</button>
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
            <div class="card-header">Resumo do Inventário</div>
            <div class="card-body">
"@

if ($totalDevices -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Dispositivos:</strong> $totalDevices</p>
                    <p><strong>Servidores:</strong> $totalServers</p>
                    <p><strong>Estações de Trabalho:</strong> $totalWorkstations</p>
                    <p><strong>Dispositivos Desativados:</strong> $disabledDevices</p>
                </div>
"@

    if ($inactiveDevices -gt 0 -and $inactiveDevices -gt ($totalDevices * 0.1)) {
        $percentInactive = [math]::Round(($inactiveDevices / $totalDevices) * 100, 1)
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$inactiveDevices dispositivos ativos sem login recente ($percentInactive% do total)</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum dispositivo encontrado ou você não tem permissão para visualizar</span>
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
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-hdd"></i>
            <h3>$totalDevices</h3>
            <p>Total de Dispositivos</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-server"></i>
            <h3>$totalServers</h3>
            <p>Servidores</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-desktop"></i>
            <h3>$totalWorkstations</h3>
            <p>Estações de Trabalho</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Tipo de Dispositivos</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="deviceTypeChart"></canvas>
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

<!-- Gráfico de sistemas operacionais -->
<div class="card mb-4">
    <div class="card-header">Distribuição de Sistemas Operacionais</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="osChart"></canvas>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Dados Detalhados de Inventário</div>
        <div>
            <span class="badge badge-primary" id="recordCount">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
"@

if ($totalDevices -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="inventoryTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>IP</th>
                        <th>Sistema Operacional</th>
                        <th>Tipo</th>
                        <th>Último Login</th>
                        <th>Status</th>
                        <th>Descrição</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($device in $inventoryData) {
        $statusBadge = if ($device.Status -eq "Ativo") {
            '<span class="badge-status badge-success">Ativo</span>'
        } else {
            '<span class="badge-status badge-danger">Inativo</span>'
        }
        
        $typeBadge = if ($device.TipoDispositivo -eq "Servidor") {
            '<span class="badge-status badge-server">Servidor</span>'
        } else {
            '<span class="badge-status badge-workstation">Estação</span>'
        }
        
        $bodyContent += @"
                    <tr data-type="$($device.TipoDispositivo.ToLower())">
                        <td>$($device.Nome)</td>
                        <td>$($device.IP)</td>
                        <td>$($device.SistemaOperacional)</td>
                        <td>$typeBadge</td>
                        <td>$($device.UltimoLogin)</td>
                        <td>$statusBadge</td>
                        <td class="truncate" title="$($device.Descricao)">$($device.Descricao)</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewDevice('$($device.Nome)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="pingDevice('$($device.IP)')"><i class="fas fa-network-wired"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhum dispositivo encontrado ou você não tem permissão para visualizar.
        </div>
"@
}

$bodyContent += @"
    </div>
</div>
"@

# Preparar dados para gráficos
$osLabels = "["
$osData = "["
$osColors = "["

$colorPalette = @(
    '#6a3094', '#9657c7', '#8244b2', '#2c1445', '#c9a6e9', 
    '#5d2683', '#4c1d6b', '#e0cdf0', '#b688df', '#7e3fa8'
)

$i = 0
foreach ($os in $osList.Keys) {
    $osLabels += "'$os',"
    $osData += "$($osList[$os]),"
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
        // Gráfico de tipos de dispositivos
        const typeCtx = document.getElementById('deviceTypeChart').getContext('2d');
        const typeChart = new Chart(typeCtx, {
            type: 'pie',
            data: {
                labels: ['Servidores', 'Estações de Trabalho'],
                datasets: [{
                    data: [$totalServers, $totalWorkstations],
                    backgroundColor: [
                        '#6a3094', // Roxo para Servidores
                        '#9657c7'  // Roxo claro para Estações
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom'
                    }
                }
            }
        });
        
        // Gráfico de status de dispositivos
        const statusCtx = document.getElementById('deviceStatusChart').getContext('2d');
        const statusChart = new Chart(statusCtx, {
            type: 'pie',
            data: {
                labels: ['Ativos', 'Inativos', 'Sem Login Recente'],
                datasets: [{
                    data: [$totalDevices - $disabledDevices - $inactiveDevices, $disabledDevices, $inactiveDevices],
                    backgroundColor: [
                        '#28a745',  // Verde para Ativos
                        '#dc3545',  // Vermelho para Inativos
                        '#ffc107'   // Amarelo para Sem Login Recente
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom'
                    }
                }
            }
        });
        
        // Gráfico de sistemas operacionais
        const osCtx = document.getElementById('osChart').getContext('2d');
        const osChart = new Chart(osCtx, {
            type: 'bar',
            data: {
                labels: $osLabels,
                datasets: [{
                    label: 'Quantidade',
                    data: $osData,
                    backgroundColor: $osColors,
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
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#inventoryTable tbody tr');
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

    // Função para filtrar dispositivos por tipo
    function filterDevices(type) {
        const rows = document.querySelectorAll('#inventoryTable tbody tr');
        let visibleCount = 0;
        
        rows.forEach(row => {
            if (type === 'all' || row.getAttribute('data-type') === type) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });
        
        // Atualizar contagem de registros visíveis
        document.getElementById('recordCount').textContent = visibleCount + ' registros';
    }
    
    // Funções para interação com dispositivos
    function viewDevice(deviceName) {
        alert('Visualizando detalhes do dispositivo: ' + deviceName);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function pingDevice(ip) {
        if (!ip) {
            alert('Endereço IP não disponível para este dispositivo');
            return;
        }
        alert('Enviando ping para o IP: ' + ip);
        // Aqui poderia implementar uma verificação de conectividade real
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Inventário" `
                      -ActiveMenu "Inventário" `
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