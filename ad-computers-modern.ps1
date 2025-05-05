<#
.SYNOPSIS
    Active Directory - Todos os Computadores (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar computadores no Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-computers.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"
$directoryPath = (Get-Item -Path ".").FullName
#$outputPath = "ad-reports\ad-computers\ad-computers-modern.html"


# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-computers\ad-computers-modern.html"

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
Write-Host "Coletando dados de computadores do Active Directory..."
try {
    $computers = Get-ADComputer -Filter { OperatingSystem -NotLike '*Windows Server*' } -Properties Name, OperatingSystem, Description, IPv4Address, Enabled, LastLogonDate

    # Coletar estatísticas
    $totalComputers = $computers.Count
    $enabledComputers = ($computers | Where-Object { $_.Enabled -eq $true }).Count
    $disabledComputers = $totalComputers - $enabledComputers

    # Agrupar por sistema operacional para estatísticas
    $osSummary = @{}
    $computers | ForEach-Object {
        $os = if ($_.OperatingSystem) { $_.OperatingSystem } else { "Desconhecido" }
        if ($osSummary.ContainsKey($os)) {
            $osSummary[$os]++
        } else {
            $osSummary[$os] = 1
        }
    }
}
catch {
    Write-Host "Erro ao coletar informações de computadores: $_" -ForegroundColor Red
    $computers = @()
    $totalComputers = 0
    $enabledComputers = 0
    $disabledComputers = 0
    $osSummary = @{}
}

# Preparar os dados para o relatório
$computerData = @()
foreach ($computer in $computers) {
    $computerData += [PSCustomObject]@{
        Nome = $computer.Name
        SistemaOperacional = $computer.OperatingSystem
        Descricao = $computer.Description
        EnderecoIP = $computer.IPv4Address
        UltimoLogin = if ($computer.LastLogonDate) { $computer.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        Status = if ($computer.Enabled) { "Ativo" } else { "Inativo" }
    }
}

# Ordenar resultados
$computerData = $computerData | Sort-Object -Property "Nome"

# Contar registros
$totalRecords = $computerData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Todos os Computadores</h1>
    <div class="header-actions">
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
            <div class="card-header">Resumo de Computadores</div>
            <div class="card-body">
                <div class="info-box">
                    <p><strong>Total de Computadores:</strong> $totalComputers</p>
                    <p><strong>Computadores Ativos:</strong> $enabledComputers</p>
                    <p><strong>Computadores Desativados:</strong> $disabledComputers</p>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Estatísticas -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card stat-card">
            <i class="fas fa-desktop"></i>
            <h3>$enabledComputers</h3>
            <p>Computadores Ativos</p>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card stat-card">
            <i class="fas fa-power-off"></i>
            <h3>$disabledComputers</h3>
            <p>Computadores Desativados</p>
        </div>
    </div>
</div>

<!-- Gráfico de Sistemas Operacionais -->
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
        <div>Dados Detalhados</div>
        <div>
            <span class="badge badge-primary">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
        <div class="table-responsive">
            <table id="computersTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Sistema Operacional</th>
                        <th>Descrição</th>
                        <th>Endereço IP</th>
                        <th>Último Login</th>
                        <th>Status</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

# Adicionar dados à tabela
foreach ($computer in $computerData) {
    $statusClass = if ($computer.Status -eq "Ativo") { "badge-active" } else { "badge-inactive" }
    $bodyContent += @"
                    <tr>
                        <td>$($computer.Nome)</td>
                        <td>$($computer.SistemaOperacional)</td>
                        <td>$($computer.Descricao)</td>
                        <td>$($computer.EnderecoIP)</td>
                        <td>$($computer.UltimoLogin)</td>
                        <td><span class="badge-status $statusClass">$($computer.Status)</span></td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewComputer('$($computer.Nome)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="pingComputer('$($computer.EnderecoIP)')"><i class="fas fa-network-wired"></i></button>
                        </td>
                    </tr>
"@
}

# Fechar o HTML
$bodyContent += @"
                </tbody>
            </table>
        </div>
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
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#computersTable tbody tr');
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
            document.querySelector('.badge.badge-primary').textContent = visibleCount + ' registros';
        });
    });

    // Funções para interação com computadores
    function viewComputer(computerName) {
        alert('Visualizando detalhes do computador: ' + computerName);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function pingComputer(ipAddress) {
        if (!ipAddress || ipAddress === "") {
            alert('Endereço IP não disponível para este computador.');
            return;
        }
        
        alert('Testando conectividade com: ' + ipAddress);
        // Aqui poderia iniciar um teste de ping
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Computadores" `
                      -ActiveMenu "Computadores" `
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