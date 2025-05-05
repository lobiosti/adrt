<#
.SYNOPSIS
    Active Directory - Controladores de Domínio (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar controladores de domínio no Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-dcs.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"
$directoryPath = (Get-Item -Path ".").FullName

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-dcs\ad-dcs-modern.html"

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
Write-Host "Coletando dados de controladores de domínio..."
try {
    $domainControllers = Get-ADDomainController -Filter * | 
                       Select-Object Site, Name, IPv4Address, OperatingSystem, OperatingSystemServicePack, 
                                     IsGlobalCatalog, IsReadOnly
    
    # Coletar estatísticas para o dashboard
    $totalDCs = $domainControllers.Count
    $globalCatalogs = ($domainControllers | Where-Object { $_.IsGlobalCatalog -eq $true }).Count
    $readOnlyDCs = ($domainControllers | Where-Object { $_.IsReadOnly -eq $true }).Count
    $totalServers = (Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' }).Count
    $nonDCServers = $totalServers - $totalDCs
}
catch {
    Write-Host "Erro ao coletar informações de controladores de domínio: $_" -ForegroundColor Red
    $domainControllers = @()
    $totalDCs = 0
    $globalCatalogs = 0
    $readOnlyDCs = 0
    $totalServers = 0
    $nonDCServers = 0
}

# Preparar os dados para o relatório
$dcData = @()
foreach ($dc in $domainControllers) {
    $dcData += [PSCustomObject]@{
        Site = $dc.Site
        Nome = $dc.Name
        EnderecoIP = $dc.IPv4Address
        SistemaOperacional = $dc.OperatingSystem
        ServicePack = $dc.OperatingSystemServicePack
        CatalogoGlobal = if ($dc.IsGlobalCatalog) { "Sim" } else { "Não" }
        SomenteLeitura = if ($dc.IsReadOnly) { "Sim" } else { "Não" }
    }
}

# Ordenar resultados
$dcData = $dcData | Sort-Object -Property "Site", "Nome"

# Contar registros
$totalRecords = $dcData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Controladores de Domínio</h1>
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
                <div class="info-box">
                    <p><strong>Total de Controladores de Domínio:</strong> $totalDCs</p>
                    <p><strong>Catálogos Globais:</strong> $globalCatalogs</p>
                    <p><strong>Controladores Somente Leitura (RODC):</strong> $readOnlyDCs</p>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Estatísticas -->
<div class="row mb-4">
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-shield-alt"></i>
            <h3>$totalDCs</h3>
            <p>Controladores de Domínio</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-globe"></i>
            <h3>$globalCatalogs</h3>
            <p>Catálogos Globais</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-server"></i>
            <h3>$totalServers</h3>
            <p>Total de Servidores</p>
        </div>
    </div>
</div>

<!-- Gráfico -->
<div class="card mb-4">
    <div class="card-header">Distribuição de Servidores</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="serverChart"></canvas>
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
            <table id="dcsTable">
                <thead>
                    <tr>
                        <th>Site</th>
                        <th>Nome</th>
                        <th>Endereço IP</th>
                        <th>Sistema Operacional</th>
                        <th>Service Pack</th>
                        <th>Catálogo Global</th>
                        <th>Somente Leitura</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

# Adicionar dados à tabela
foreach ($dc in $dcData) {
    $catalogoStatus = if ($dc.CatalogoGlobal -eq "Sim") {
        '<span class="badge-status badge-success">Sim</span>'
    } else {
        '<span class="badge-status badge-danger">Não</span>'
    }
    
    $rodcStatus = if ($dc.SomenteLeitura -eq "Sim") {
        '<span class="badge-status badge-warning">Sim</span>'
    } else {
        '<span class="badge-status badge-success">Não</span>'
    }
    
    $bodyContent += @"
                    <tr>
                        <td>$($dc.Site)</td>
                        <td>$($dc.Nome)</td>
                        <td>$($dc.EnderecoIP)</td>
                        <td>$($dc.SistemaOperacional)</td>
                        <td>$($dc.ServicePack)</td>
                        <td>$catalogoStatus</td>
                        <td>$rodcStatus</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewDC('$($dc.Nome)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="pingDC('$($dc.EnderecoIP)')"><i class="fas fa-network-wired"></i></button>
                        </td>
                    </tr>
"@
}

$bodyContent += @"
                </tbody>
            </table>
        </div>
    </div>
</div>
"@

# Script específico para esta página
$extraScripts = @"
<script>
    // Inicializar o gráfico
    document.addEventListener('DOMContentLoaded', function() {
        const ctx = document.getElementById('serverChart').getContext('2d');
        const serverChart = new Chart(ctx, {
            type: 'pie',
            data: {
                labels: ['Controladores de Domínio', 'Outros Servidores'],
                datasets: [{
                    data: [$totalDCs, $nonDCServers],
                    backgroundColor: [
                        '#6a3094', // Roxo para DCs
                        '#9657c7'  // Roxo claro para outros servidores
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                    }
                }
            }
        });
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#dcsTable tbody tr');
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

    // Funções para interação com DCs
    function viewDC(dcName) {
        alert('Visualizando detalhes do DC: ' + dcName);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function pingDC(ipAddress) {
        alert('Testando conectividade com: ' + ipAddress);
        // Aqui poderia iniciar um teste de ping
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
                      -Title "Active Directory Report Tool - Controladores de Domínio" `
                      -ActiveMenu "Controladores de Domínio" `
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