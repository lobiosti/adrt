<#
.SYNOPSIS
    Active Directory - Políticas de Grupo (GPOs) (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar e visualizar políticas de grupo (GPOs) do Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-gpos.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-gpos\ad-gpos-modern.html"
$gpoDetailDir = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-gpos\gpos-html"
$gpoZipFile = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-gpos\gpos-html.zip"

# Criar diretórios se não existirem
$outputDir = Split-Path -Path $outputPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

if (-not (Test-Path -Path $gpoDetailDir)) {
    New-Item -ItemType Directory -Path $gpoDetailDir -Force | Out-Null
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

# Importar módulos necessários
Import-Module ActiveDirectory
Import-Module GroupPolicy

# Coletar dados
Write-Host "Coletando dados de GPOs do Active Directory..."
try {
    # Obter todas as GPOs
    $gpos = Get-GPO -All | 
            Select-Object DisplayName, ID, Owner, CreationTime, ModificationTime, Description, 
                      @{Name="Computer"; Expression={$_.Computer.Enabled}},
                      @{Name="User"; Expression={$_.User.Enabled}}
    
    # Coletar estatísticas para o dashboard
    $totalGPOs = $gpos.Count
    $computerEnabled = ($gpos | Where-Object { $_.Computer -eq $true }).Count
    $userEnabled = ($gpos | Where-Object { $_.User -eq $true }).Count
    $bothEnabled = ($gpos | Where-Object { $_.Computer -eq $true -and $_.User -eq $true }).Count
    $noneEnabled = ($gpos | Where-Object { $_.Computer -eq $false -and $_.User -eq $false }).Count
    
    # Descobrir datas
    $recentGPOs = ($gpos | Where-Object { $_.ModificationTime -gt (Get-Date).AddDays(-30) }).Count
    $oldGPOs = ($gpos | Where-Object { $_.ModificationTime -lt (Get-Date).AddYears(-1) }).Count
    
    # Gerar relatórios HTML detalhados para cada GPO
    Write-Host "Gerando relatórios detalhados para cada GPO..."
    foreach ($gpo in $gpos) {
        try {
            # Gerar relatório HTML para cada GPO
            $gpoHtmlPath = Join-Path -Path $gpoDetailDir -ChildPath "$($gpo.DisplayName).htm"
            $gpo | Get-GPOReport -ReportType HTML | Out-File -FilePath $gpoHtmlPath -Encoding utf8 -Force
        }
        catch {
            Write-Host "Erro ao gerar relatório para GPO '$($gpo.DisplayName)': $_" -ForegroundColor Yellow
        }
    }
    
    # Compactar relatórios para facilitar o download
    try {
        Compress-Archive -Path "$gpoDetailDir\*" -DestinationPath $gpoZipFile -Force
        Write-Host "Relatórios HTML compactados em: $gpoZipFile" -ForegroundColor Green
    }
    catch {
        Write-Host "Erro ao compactar relatórios: $_" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Erro ao coletar informações de GPOs: $_" -ForegroundColor Red
    $gpos = @()
    $totalGPOs = 0
    $computerEnabled = 0
    $userEnabled = 0
    $bothEnabled = 0
    $noneEnabled = 0
    $recentGPOs = 0
    $oldGPOs = 0
}

# Preparar os dados para o relatório
$gpoData = @()
foreach ($gpo in $gpos) {
    # Definir status de modificação
    $modificationStatus = "Normal"
    if ($gpo.ModificationTime -gt (Get-Date).AddDays(-30)) {
        $modificationStatus = "Recente"
    } elseif ($gpo.ModificationTime -lt (Get-Date).AddYears(-1)) {
        $modificationStatus = "Antigo"
    }
    
    $gpoData += [PSCustomObject]@{
        Nome = $gpo.DisplayName
        ID = $gpo.ID
        Proprietario = $gpo.Owner
        Criado = if ($gpo.CreationTime) { $gpo.CreationTime.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        UltimaModificacao = if ($gpo.ModificationTime) { $gpo.ModificationTime.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Descricao = $gpo.Description
        ConfigComputador = $gpo.Computer
        ConfigUsuario = $gpo.User
        ArquivoHTML = "$($gpo.DisplayName).htm"
        StatusModificacao = $modificationStatus
    }
}

# Ordenar resultados
$gpoData = $gpoData | Sort-Object -Property "Nome"

# Contar registros
$totalRecords = $gpoData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Políticas de Grupo (GPOs)</h1>
    <div class="header-actions">
        <button onclick="filterGPOs('all')"><i class="fas fa-sync"></i> Mostrar Todas</button>
        <button onclick="filterGPOs('recent')"><i class="fas fa-history"></i> Modificadas Recentemente</button>
        <button onclick="downloadGPOReports()"><i class="fas fa-file-archive"></i> Baixar Relatórios</button>
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
            <div class="card-header">Resumo de Políticas de Grupo</div>
            <div class="card-body">
"@

if ($totalGPOs -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de GPOs:</strong> $totalGPOs</p>
                    <p><strong>GPOs com Configurações de Computador:</strong> $computerEnabled</p>
                    <p><strong>GPOs com Configurações de Usuário:</strong> $userEnabled</p>
                    <p><strong>GPOs com Ambas Configurações:</strong> $bothEnabled</p>
                    <p><strong>GPOs Modificadas nos Últimos 30 Dias:</strong> $recentGPOs</p>
                </div>
"@

    if ($noneEnabled -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$noneEnabled GPOs sem configurações de computador ou usuário ativas</span>
                </div>
"@
    }
    
    if ($oldGPOs -gt ($totalGPOs * 0.7)) {
        $oldPercent = [math]::Round(($oldGPOs / $totalGPOs) * 100, 1)
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$oldPercent% das GPOs não foram atualizadas há mais de 1 ano</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhuma GPO encontrada ou você não tem permissão para visualizar</span>
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
            <i class="fas fa-cogs"></i>
            <h3>$totalGPOs</h3>
            <p>Total de GPOs</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-desktop"></i>
            <h3>$computerEnabled</h3>
            <p>Configurações de Computador</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-user-cog"></i>
            <h3>$userEnabled</h3>
            <p>Configurações de Usuário</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Tipos de Configuração</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="configTypeChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Atualização de GPOs</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="modificationChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Políticas de Grupo Detalhadas</div>
        <div>
            <span class="badge badge-primary" id="recordCount">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
"@

if ($totalGPOs -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="gpoTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Descrição</th>
                        <th>Proprietário</th>
                        <th>Criado</th>
                        <th>Última Modificação</th>
                        <th>Config. Computador</th>
                        <th>Config. Usuário</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($gpo in $gpoData) {
        $computerStatusBadge = if ($gpo.ConfigComputador) {
            '<span class="badge-status badge-success">Ativo</span>'
        } else {
            '<span class="badge-status badge-danger">Inativo</span>'
        }
        
        $userStatusBadge = if ($gpo.ConfigUsuario) {
            '<span class="badge-status badge-success">Ativo</span>'
        } else {
            '<span class="badge-status badge-danger">Inativo</span>'
        }
        
        $modificationClass = "status-$($gpo.StatusModificacao.ToLower())"
        
        $bodyContent += @"
                    <tr data-status="$($gpo.StatusModificacao.ToLower())">
                        <td>$($gpo.Nome)</td>
                        <td class="truncate" title="$($gpo.Descricao)">$($gpo.Descricao)</td>
                        <td>$($gpo.Proprietario)</td>
                        <td>$($gpo.Criado)</td>
                        <td class="$modificationClass">$($gpo.UltimaModificacao)</td>
                        <td>$computerStatusBadge</td>
                        <td>$userStatusBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewGPODetails('$($gpo.ArquivoHTML)', '$($gpo.Nome)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="viewGPOSettings('$($gpo.ID)')"><i class="fas fa-list-check"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhuma GPO encontrada ou você não tem permissão para visualizar.
        </div>
"@
}

$bodyContent += @"
    </div>
</div>

<!-- Modal para visualização detalhada da GPO -->
<div id="gpoDetailModal" class="gpo-modal">
    <div class="gpo-modal-content">
        <div class="gpo-modal-header">
            <h3 id="gpoModalTitle">Detalhes da GPO</h3>
            <span class="close-modal" onclick="closeGPOModal()">&times;</span>
        </div>
        <iframe id="gpoModalIframe" class="gpo-modal-iframe" src=""></iframe>
    </div>
</div>
"@

# Script específico para esta página
$extraScripts = @"
<style>
    /* Estilos para o modal */
    .gpo-modal {
        display: none;
        position: fixed;
        z-index: 2000;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        overflow: auto;
        background-color: rgba(0, 0, 0, 0.5);
    }
    
    .gpo-modal-content {
        background-color: #fefefe;
        margin: 2% auto;
        padding: 20px;
        border: 1px solid #888;
        border-radius: 10px;
        width: 90%;
        height: 90%;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
    }
    
    .gpo-modal-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding-bottom: 10px;
        border-bottom: 1px solid #eee;
        margin-bottom: 15px;
    }
    
    .gpo-modal-iframe {
        width: 100%;
        height: calc(100% - 80px);
        border: none;
        border-radius: 5px;
    }
    
    .close-modal {
        color: var(--lobios-primary);
        font-size: 28px;
        font-weight: bold;
        cursor: pointer;
    }
    
    /* Estilo para destacar status de modificação */
    .status-recente {
        color: #d63384;
        font-weight: bold;
    }
    
    .status-antigo {
        color: #6c757d;
    }
</style>

<script>
    // Inicializar gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de tipos de configuração
        const configCtx = document.getElementById('configTypeChart').getContext('2d');
        const configChart = new Chart(configCtx, {
            type: 'pie',
            data: {
                labels: ['Ambas Configurações', 'Apenas Computador', 'Apenas Usuário', 'Nenhuma Configuração'],
                datasets: [{
                    data: [$bothEnabled, $($computerEnabled - $bothEnabled), $($userEnabled - $bothEnabled), $noneEnabled],
                    backgroundColor: [
                        '#6a3094', // Roxo para Ambas
                        '#3f51b5', // Azul para Computador
                        '#2196f3', // Azul claro para Usuário
                        '#dc3545'  // Vermelho para Nenhuma
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
        
        // Gráfico de atualização de GPOs
        const modCtx = document.getElementById('modificationChart').getContext('2d');
        const modChart = new Chart(modCtx, {
            type: 'pie',
            data: {
                labels: ['Recentes (< 30 dias)', 'Normais (30d - 1a)', 'Antigas (> 1 ano)'],
                datasets: [{
                    data: [$recentGPOs, $($totalGPOs - $recentGPOs - $oldGPOs), $oldGPOs],
                    backgroundColor: [
                        '#28a745',  // Verde para Recentes
                        '#fd7e14',  // Laranja para Normais
                        '#6c757d'   // Cinza para Antigas
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
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#gpoTable tbody tr');
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

    // Função para filtrar GPOs por status de modificação
    function filterGPOs(status) {
        const rows = document.querySelectorAll('#gpoTable tbody tr');
        let visibleCount = 0;
        
        rows.forEach(row => {
            if (status === 'all' || row.getAttribute('data-status') === status) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });
        
        // Atualizar contagem de registros visíveis
        document.getElementById('recordCount').textContent = visibleCount + ' registros';
    }
    
    // Funções para visualização de GPOs
    function viewGPODetails(htmlFile, gpoName) {
        const modal = document.getElementById('gpoDetailModal');
        const iframe = document.getElementById('gpoModalIframe');
        const title = document.getElementById('gpoModalTitle');
        
        title.textContent = 'Detalhes da GPO: ' + gpoName;
        iframe.src = 'gpos-html/' + htmlFile;
        modal.style.display = 'block';
    }
    
    function closeGPOModal() {
        const modal = document.getElementById('gpoDetailModal');
        modal.style.display = 'none';
    }
    
    function viewGPOSettings(gpoId) {
        alert('Visualizando configurações da GPO com ID: ' + gpoId);
        // Aqui poderia abrir o editor de GPO local, mas isso requer integração com o sistema operacional
    }
    
    // Função para baixar o arquivo ZIP com relatórios de GPOs
    function downloadGPOReports() {
        window.location.href = 'gpos-html.zip';
    }
    
    // Função para exportar para CSV
    function exportToCsv() {
        alert('Exportando dados para CSV...');
        // Implementação da exportação CSV
    }
    
    // Fechar o modal se o usuário clicar fora dele
    window.onclick = function(event) {
        const modal = document.getElementById('gpoDetailModal');
        if (event.target == modal) {
            modal.style.display = 'none';
        }
    };
    
    // Adicionar tooltip para elementos truncados
    document.querySelectorAll('.truncate').forEach(function(element) {
        element.addEventListener('mouseover', function() {
            if (this.offsetWidth < this.scrollWidth) {
                this.title = this.textContent;
            }
        });
    });
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Políticas de Grupo" `
                      -ActiveMenu "GPOs" `
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
    Write-Host "Detalhes de GPOs individuais salvos em: $gpoDetailDir" -ForegroundColor Green
    Write-Host "Arquivo ZIP com todos os relatórios: $gpoZipFile" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao salvar o relatório: $_" -ForegroundColor Red
}

# Abrir o relatório no navegador
Start-Process $outputPath