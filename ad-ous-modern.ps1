<#
.SYNOPSIS
    Active Directory - Todas as Unidades Organizacionais (OUs) (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar todas as Unidades Organizacionais (OUs) do Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-ous.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-ous\ad-ous-modern.html"

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
Write-Host "Coletando dados de Unidades Organizacionais (OUs)..."
try {
    # Obter todas as OUs
    $ous = Get-ADOrganizationalUnit -Filter * -Properties Name, Description, Created, Modified, ProtectedFromAccidentalDeletion, DistinguishedName
    
    # Contagem de OUs por nível de profundidade
    $ousByDepth = @{}
    foreach ($ou in $ous) {
        # Calcular profundidade da OU baseado no DistinguishedName
        $depth = ($ou.DistinguishedName -split ',').Count - ($ou.DistinguishedName -split ',DC=').Count
        if ($ousByDepth.ContainsKey($depth)) {
            $ousByDepth[$depth]++
        } else {
            $ousByDepth[$depth] = 1
        }
    }
    
    # Contagem total
    $totalOUs = $ous.Count
    
    # Calcular quantas OUs estão protegidas contra exclusão acidental
    $protectedOUs = ($ous | Where-Object { $_.ProtectedFromAccidentalDeletion -eq $true }).Count
    $protectedPercentage = if ($totalOUs -gt 0) { [math]::Round(($protectedOUs / $totalOUs) * 100, 1) } else { 0 }
    
    # Obter estatísticas adicionais do AD para contexto
    $totalUsers = (Get-ADUser -Filter *).Count
    $totalComputers = (Get-ADComputer -Filter *).Count
}
catch {
    Write-Host "Erro ao coletar informações de OUs: $_" -ForegroundColor Red
    $ous = @()
    $totalOUs = 0
    $protectedOUs = 0
    $protectedPercentage = 0
    $totalUsers = 0
    $totalComputers = 0
    $ousByDepth = @{}
}

# Preparar os dados para o relatório
$ouData = @()
foreach ($ou in $ous) {
    # Extrair o nome da OU do DN
    $ouName = ($ou.DistinguishedName -split ',')[0] -replace 'OU=', ''
    
    # Calcular o nível hierárquico
    $level = ($ou.DistinguishedName -split ',').Count - ($ou.DistinguishedName -split ',DC=').Count
    
   # Criar caminho hierárquico para exibição
    $path = ""
    $parts = $ou.DistinguishedName -split ','
    $dcCount = ($parts | Where-Object { $_ -like "DC=*" }).Count
    $nonDcCount = $parts.Count - $dcCount

    for ($i = 0; $i -lt $nonDcCount; $i++) {
        if ($parts[$i] -like "OU=*") {
            $name = $parts[$i] -replace 'OU=', ''
            $path = if ($path -eq "") { $name } else { "$name > $path" }
        }
    }
    
    $ouData += [PSCustomObject]@{
        Nome = $ouName
        Descricao = $ou.Description
        NivelHierarquico = $level
        CaminhoCompleto = $path
        DataCriacao = if ($ou.Created) { $ou.Created.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        UltimaModificacao = if ($ou.Modified) { $ou.Modified.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Protegida = $ou.ProtectedFromAccidentalDeletion
        DN = $ou.DistinguishedName
    }
}

# Ordenar resultados hierarquicamente
$ouData = $ouData | Sort-Object -Property NivelHierarquico, CaminhoCompleto

# Contar registros
$totalRecords = $ouData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Unidades Organizacionais (OUs)</h1>
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
            <div class="card-header">Resumo da Estrutura</div>
            <div class="card-body">
"@

if ($totalOUs -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de OUs:</strong> $totalOUs</p>
                    <p><strong>OUs protegidas contra exclusão acidental:</strong> $protectedOUs ($protectedPercentage%)</p>
                </div>
"@

    if ($protectedPercentage -lt 50) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>Apenas $protectedPercentage% das OUs estão protegidas contra exclusão acidental.</span>
                </div>
"@
    } elseif ($protectedPercentage -ge 90) {
        $bodyContent += @"
                <div class="warning-item warning-green">
                    <i class="fas fa-check-circle"></i>
                    <span>Excelente: $protectedPercentage% das OUs estão protegidas contra exclusão acidental.</span>
                </div>
"@
    }
    
    if ($totalOUs -gt 50) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Grande número de OUs ($totalOUs) pode indicar uma estrutura complexa. Considere revisar a hierarquia.</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhuma OU encontrada ou você não tem permissão para visualizar.</span>
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
            <i class="fas fa-sitemap"></i>
            <h3>$totalOUs</h3>
            <p>Unidades Organizacionais</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-users"></i>
            <h3>$totalUsers</h3>
            <p>Total de Usuários</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-desktop"></i>
            <h3>$totalComputers</h3>
            <p>Total de Computadores</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribuição de OUs por Nível Hierárquico</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="ouLevelChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Proteção de OUs</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="protectionChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Estrutura Hierárquica de OUs</div>
        <div>
            <span class="badge badge-primary">$totalRecords registros</span>
        </div>
    </div>
    <div class="card-body">
"@

if ($totalOUs -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Descrição</th>
                        <th>Caminho</th>
                        <th>Data de Criação</th>
                        <th>Nível</th>
                        <th>Protegida</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($ou in $ouData) {
        $protectionBadge = if ($ou.Protegida) {
            '<span class="badge-status badge-success">Sim</span>'
        } else {
            '<span class="badge-status badge-danger">Não</span>'
        }
        
        $ouLevelClass = "ou-level-$($ou.NivelHierarquico)"
        
        $bodyContent += @"
                    <tr>
                        <td class="$ouLevelClass"><i class="fas fa-folder"></i> $($ou.Nome)</td>
                        <td>$($ou.Descricao)</td>
                        <td>$($ou.CaminhoCompleto)</td>
                        <td>$($ou.DataCriacao)</td>
                        <td>$($ou.NivelHierarquico)</td>
                        <td>$protectionBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewOU('$($ou.DN)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="editOU('$($ou.DN)')"><i class="fas fa-edit"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhuma Unidade Organizacional encontrada neste domínio ou você não tem permissão para visualizar.
        </div>
"@
}

$bodyContent += @"
    </div>
</div>
"@

# Preparar dados para gráficos
$ouLevelLabels = "["
$ouLevelData = "["

foreach ($depth in $ousByDepth.Keys | Sort-Object) {
    $ouLevelLabels += "'Nível $depth', "
    $ouLevelData += "$($ousByDepth[$depth]), "
}

$ouLevelLabels = $ouLevelLabels.TrimEnd(',') + "]"
$ouLevelData = $ouLevelData.TrimEnd(',') + "]"

# Script específico para esta página
$extraScripts = @"
<style>
    /* Exibição de árvore hierárquica */
    .ou-level-1 { padding-left: 10px; }
    .ou-level-2 { padding-left: 30px; }
    .ou-level-3 { padding-left: 50px; }
    .ou-level-4 { padding-left: 70px; }
    .ou-level-5 { padding-left: 90px; }
    .ou-level-6 { padding-left: 110px; }
</style>

<script>
    // Inicializar os gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de distribuição de OUs por nível
        const levelCtx = document.getElementById('ouLevelChart').getContext('2d');
        const ouLevelChart = new Chart(levelCtx, {
            type: 'bar',
            data: {
                labels: $ouLevelLabels,
                datasets: [{
                    label: 'Quantidade de OUs',
                    data: $ouLevelData,
                    backgroundColor: '#6a3094',
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
        
        // Gráfico de proteção de OUs
        const protectionCtx = document.getElementById('protectionChart').getContext('2d');
        const protectionChart = new Chart(protectionCtx, {
            type: 'pie',
            data: {
                labels: ['Protegidas', 'Não Protegidas'],
                datasets: [{
                    data: [$protectedOUs, $($totalOUs - $protectedOUs)],
                    backgroundColor: [
                        '#28a745', // Verde para protegidas
                        '#dc3545'  // Vermelho para não protegidas
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
    });

    // Funções para interação com OUs
    function viewOU(dn) {
        alert('Visualizando detalhes da OU: ' + dn);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function editOU(dn) {
        alert('Editando OU: ' + dn);
        // Aqui poderia abrir um formulário de edição
    }
    
    // Função para exportar para CSV
    function exportToCsv() {
        alert('Exportando dados para CSV...');
        // Implementação da exportação CSV
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Unidades Organizacionais" `
                      -ActiveMenu "OUs" `
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