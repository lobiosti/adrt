<#
.SYNOPSIS
    Active Directory - Grupos (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar todos os grupos do Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-groups.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-groups\ad-groups-modern.html"

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
Write-Host "Coletando dados de grupos..."
try {
    # Obter todos os grupos com propriedades estendidas
    $groups = Get-ADGroup -Filter {name -like "*"} -Properties Name, Description, GroupCategory, GroupScope, Created, Modified, info, ManagedBy
    
    # Coletar estatísticas para o dashboard
    $totalGroups = $groups.Count
    $securityGroups = ($groups | Where-Object { $_.GroupCategory -eq "Security" }).Count
    $distributionGroups = ($groups | Where-Object { $_.GroupCategory -eq "Distribution" }).Count
    
    # Categorias de escopo
    $domainLocalGroups = ($groups | Where-Object { $_.GroupScope -eq "DomainLocal" }).Count
    $globalGroups = ($groups | Where-Object { $_.GroupScope -eq "Global" }).Count
    $universalGroups = ($groups | Where-Object { $_.GroupScope -eq "Universal" }).Count
    
    # Grupos com gerenciador definido
    $managedGroups = ($groups | Where-Object { $_.ManagedBy -ne $null }).Count
    
    # Identificar grupos padrão do sistema
    $systemGroups = ($groups | Where-Object { 
        $_.Name -in @("Domain Users", "Domain Computers", "Domain Admins", 
                      "Enterprise Admins", "Schema Admins", "Administrators", 
                      "Domain Controllers", "Cert Publishers", "Group Policy Creator Owners")
    }).Count
    
    # Grupos personalizados (não padrão)
    $customGroups = $totalGroups - $systemGroups
}
catch {
    Write-Host "Erro ao coletar informações de grupos: $_" -ForegroundColor Red
    $groups = @()
    $totalGroups = 0
    $securityGroups = 0
    $distributionGroups = 0
    $domainLocalGroups = 0
    $globalGroups = 0
    $universalGroups = 0
    $managedGroups = 0
    $systemGroups = 0
    $customGroups = 0
}

# Preparar os dados para o relatório
$groupData = @()
foreach ($group in $groups) {
    # Tentar obter o nome do gerenciador, se existir
    $managerName = ""
    if ($group.ManagedBy) {
        try {
            $manager = Get-ADObject -Identity $group.ManagedBy -Properties displayName
            $managerName = $manager.displayName
        } catch {
            # Ignorar erro e continuar
            $managerName = "Não disponível"
        }
    }
    
    $groupData += [PSCustomObject]@{
        Nome = $group.Name
        Descricao = $group.Description
        Categoria = $group.GroupCategory
        Escopo = $group.GroupScope
        Criado = if ($group.Created) { $group.Created.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Modificado = if ($group.Modified) { $group.Modified.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Gerenciador = $managerName
        Info = $group.info
        Sistema = if ($group.Name -in @("Domain Users", "Domain Computers", "Domain Admins", 
                                       "Enterprise Admins", "Schema Admins", "Administrators", 
                                       "Domain Controllers", "Cert Publishers", "Group Policy Creator Owners")) 
                 { $true } else { $false }
    }
}

# Ordenar resultados
$groupData = $groupData | Sort-Object -Property "Nome"

# Contar registros
$totalRecords = $groupData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Grupos</h1>
    <div class="header-actions">
        <button onclick="filterGroups('all')"><i class="fas fa-sync"></i> Mostrar Todos</button>
        <button onclick="filterGroups('security')"><i class="fas fa-shield-alt"></i> Segurança</button>
        <button onclick="filterGroups('distribution')"><i class="fas fa-envelope"></i> Distribuição</button>
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
            <div class="card-header">Resumo de Grupos</div>
            <div class="card-body">
"@

if ($totalGroups -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Grupos:</strong> $totalGroups</p>
                    <p><strong>Grupos de Segurança:</strong> $securityGroups</p>
                    <p><strong>Grupos de Distribuição:</strong> $distributionGroups</p>
                    <p><strong>Grupos com Gerenciador:</strong> $managedGroups</p>
                </div>
"@

    if ($managedGroups -lt ($totalGroups * 0.3)) {
        $percentManaged = [math]::Round(($managedGroups / $totalGroups) * 100, 1)
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>Apenas $percentManaged% dos grupos têm um gerenciador definido</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum grupo encontrado ou você não tem permissão para visualizar</span>
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
            <i class="fas fa-users-cog"></i>
            <h3>$totalGroups</h3>
            <p>Total de Grupos</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-shield-alt"></i>
            <h3>$securityGroups</h3>
            <p>Grupos de Segurança</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-envelope"></i>
            <h3>$distributionGroups</h3>
            <p>Grupos de Distribuição</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Tipos de Grupos</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="groupTypeChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Escopo de Grupos</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="groupScopeChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Gráfico adicional -->
<div class="card mb-4">
    <div class="card-header">Distribuição de Grupos</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="groupDistributionChart"></canvas>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Dados Detalhados</div>
        <div>
            <span class="badge badge-primary" id="recordCount">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
"@

if ($totalGroups -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="groupsTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Descrição</th>
                        <th>Categoria</th>
                        <th>Escopo</th>
                        <th>Criado</th>
                        <th>Modificado</th>
                        <th>Gerenciador</th>
                        <th>Tipo</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($group in $groupData) {
        $categoryBadge = switch ($group.Categoria) {
            "Security" { '<span class="badge-status badge-success" data-category="security">Segurança</span>' }
            "Distribution" { '<span class="badge-status badge-info" data-category="distribution">Distribuição</span>' }
            default { '<span class="badge-status badge-warning">Desconhecido</span>' }
        }
        
        $scopeBadge = switch ($group.Escopo) {
            "DomainLocal" { '<span class="badge-status badge-warning">Local de Domínio</span>' }
            "Global" { '<span class="badge-status badge-success">Global</span>' }
            "Universal" { '<span class="badge-status badge-info">Universal</span>' }
            default { '<span class="badge-status badge-warning">Desconhecido</span>' }
        }
        
        $typeBadge = if ($group.Sistema) {
            '<span class="badge-status badge-system">Sistema</span>'
        } else {
            '<span class="badge-status badge-custom">Personalizado</span>'
        }
        
        $bodyContent += @"
                    <tr data-category="$(if ($group.Categoria -eq 'Security') { 'security' } else { 'distribution' })" data-system="$(if ($group.Sistema) { 'true' } else { 'false' })">
                        <td>$($group.Nome)</td>
                        <td class="truncate" title="$($group.Descricao)">$($group.Descricao)</td>
                        <td>$categoryBadge</td>
                        <td>$scopeBadge</td>
                        <td>$($group.Criado)</td>
                        <td>$($group.Modificado)</td>
                        <td>$($group.Gerenciador)</td>
                        <td>$typeBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewGroup('$($group.Nome)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="editGroup('$($group.Nome)')"><i class="fas fa-edit"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhum grupo encontrado ou você não tem permissão para visualizar.
        </div>
"@
}

$bodyContent += @"
    </div>
</div>
"@

# Script específico para esta página
$extraScripts = @"
<script>
    // Inicializar gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de tipos de grupos
        const typeCtx = document.getElementById('groupTypeChart').getContext('2d');
        const typeChart = new Chart(typeCtx, {
            type: 'pie',
            data: {
                labels: ['Segurança', 'Distribuição'],
                datasets: [{
                    data: [$securityGroups, $distributionGroups],
                    backgroundColor: [
                        '#6a3094', // Roxo para Segurança
                        '#9657c7'  // Roxo claro para Distribuição
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
        
        // Gráfico de escopo
        const scopeCtx = document.getElementById('groupScopeChart').getContext('2d');
        const scopeChart = new Chart(scopeCtx, {
            type: 'pie',
            data: {
                labels: ['Local de Domínio', 'Global', 'Universal'],
                datasets: [{
                    data: [$domainLocalGroups, $globalGroups, $universalGroups],
                    backgroundColor: [
                        '#6a3094',  // Roxo para Local de Domínio
                        '#9657c7',  // Roxo claro para Global
                        '#c9a6e9'   // Roxo mais claro para Universal
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
        
        // Gráfico de distribuição
        const distCtx = document.getElementById('groupDistributionChart').getContext('2d');
        const distChart = new Chart(distCtx, {
            type: 'bar',
            data: {
                labels: ['Total de Grupos', 'Sistema', 'Personalizados', 'Com Gerenciador', 'Sem Gerenciador'],
                datasets: [{
                    label: 'Número de Grupos',
                    data: [$totalGroups, $systemGroups, $customGroups, $managedGroups, $totalGroups - $managedGroups],
                    backgroundColor: [
                        '#6a3094',  // Total
                        '#3f51b5',  // Sistema
                        '#0288d1',  // Personalizados
                        '#28a745',  // Com Gerenciador
                        '#dc3545'   // Sem Gerenciador
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
            const rows = document.querySelectorAll('#groupsTable tbody tr');
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

    // Função para filtrar grupos por categoria
    function filterGroups(category) {
        const rows = document.querySelectorAll('#groupsTable tbody tr');
        let visibleCount = 0;
        
        rows.forEach(row => {
            if (category === 'all' || row.getAttribute('data-category') === category) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });
        
        // Atualizar contagem de registros visíveis
        document.getElementById('recordCount').textContent = visibleCount + ' registros';
    }
    
    // Funções para interação com grupos
    function viewGroup(groupName) {
        alert('Visualizando detalhes do grupo: ' + groupName);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function editGroup(groupName) {
        alert('Editando grupo: ' + groupName);
        // Aqui poderia abrir um formulário de edição
    }
    
    // Função para exportar para CSV
    function exportToCsv() {
        alert('Exportando dados para CSV...');
        // Implementação da exportação CSV
    }
    
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
                      -Title "Active Directory Report Tool - Grupos" `
                      -ActiveMenu "Grupos" `
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