<#
.SYNOPSIS
    Active Directory - Membros de Grupos (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para Active Directory - Lista detalhada de membros de grupos
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-membergroups.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-membergroups\ad-membergroups-modern.html"

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
Write-Host "Coletando dados de grupos e membros..."
try {
    # Obter todos os grupos
    $memberGroups = Get-ADGroup -Filter {Name -like "*"} -Properties Name, Description, GroupCategory, GroupScope, Members, MemberOf
    
    # Coletar estatísticas para o dashboard
    $totalGroups = $memberGroups.Count
    $securityGroups = ($memberGroups | Where-Object { $_.GroupCategory -eq "Security" }).Count
    $distributionGroups = ($memberGroups | Where-Object { $_.GroupCategory -eq "Distribution" }).Count
    
    # Categorias de escopo
    $domainLocalGroups = ($memberGroups | Where-Object { $_.GroupScope -eq "DomainLocal" }).Count
    $globalGroups = ($memberGroups | Where-Object { $_.GroupScope -eq "Global" }).Count
    $universalGroups = ($memberGroups | Where-Object { $_.GroupScope -eq "Universal" }).Count
    
    # Grupos com mais membros
    $groupsWithMostMembers = $memberGroups | 
                            Select-Object Name, @{Name="MemberCount"; Expression={($_.Members | Measure-Object).Count}} |
                            Sort-Object -Property MemberCount -Descending |
                            Select-Object -First 10
                            
    # Grupos vazios (sem membros)
    $emptyGroups = ($memberGroups | Where-Object { ($_.Members | Measure-Object).Count -eq 0 }).Count
}
catch {
    Write-Host "Erro ao coletar informações de grupos: $_" -ForegroundColor Red
    $memberGroups = @()
    $totalGroups = 0
    $securityGroups = 0
    $distributionGroups = 0
    $domainLocalGroups = 0
    $globalGroups = 0
    $universalGroups = 0
    $groupsWithMostMembers = @()
    $emptyGroups = 0
}

# Preparar os dados para o relatório
$groupData = @()
foreach ($group in $memberGroups) {
    # Contar membros
    $memberCount = ($group.Members | Measure-Object).Count
    
    # Converter MemberOf em uma lista formatada
    $memberOfList = ""
    if ($group.MemberOf) {
        $memberOfNames = @()
        foreach ($memberOfDN in $group.MemberOf) {
            try {
                $memberOfGroup = Get-ADGroup -Identity $memberOfDN -Properties Name
                $memberOfNames += $memberOfGroup.Name
            } catch {
                # Ignorar erro e continuar
            }
        }
        $memberOfList = $memberOfNames -join ", "
    }
    
    # Converter Members em uma lista formatada (limitada a 10 para evitar overflow)
    $membersList = ""
    if ($group.Members) {
        $memberNames = @()
        $count = 0
        foreach ($memberDN in $group.Members) {
            if ($count -ge 10) {
                $memberNames += "... e mais $(($group.Members | Measure-Object).Count - 10) membros"
                break
            }
            
            try {
                $member = Get-ADObject -Identity $memberDN -Properties Name
                $memberNames += $member.Name
                $count++
            } catch {
                # Ignorar erro e continuar
            }
        }
        $membersList = $memberNames -join ", "
    }
    
    $groupData += [PSCustomObject]@{
        Nome = $group.Name
        Descrição = $group.Description
        Categoria = $group.GroupCategory
        Escopo = $group.GroupScope
        Membros = $membersList
        ContadorMembros = $memberCount
        MembroDe = $memberOfList
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
    <h1>Active Directory - Membros de Grupos</h1>
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
            <div class="card-header">Resumo de Grupos</div>
            <div class="card-body">
"@

if ($totalGroups -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Grupos:</strong> $totalGroups</p>
                    <p><strong>Grupos de Segurança:</strong> $securityGroups</p>
                    <p><strong>Grupos de Distribuição:</strong> $distributionGroups</p>
                    <p><strong>Grupos Vazios (sem membros):</strong> $emptyGroups</p>
                </div>
"@

        if ($emptyGroups -gt 0 -and $emptyGroups -gt ($totalGroups * 0.2)) {
            $percentEmpty = [math]::Round(($emptyGroups / $totalGroups) * 100, 1)
            $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$emptyGroups grupos vazios ($percentEmpty% do total)</span>
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

<!-- Grupos com mais membros -->
<div class="card mb-4">
    <div class="card-header">Top 10 Grupos com Mais Membros</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="memberCountChart"></canvas>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Dados Detalhados</div>
        <div>
            <span class="badge badge-primary">$totalRecords registros</span>
        </div>
    </div>
    <div class="card-body">
"@

if ($totalGroups -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Descrição</th>
                        <th>Categoria</th>
                        <th>Escopo</th>
                        <th>Qtd. Membros</th>
                        <th>Membros</th>
                        <th>Membro De</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($group in $groupData) {
        $categoryBadge = switch ($group.Categoria) {
            "Security" { '<span class="badge-status badge-success">Segurança</span>' }
            "Distribution" { '<span class="badge-status badge-info">Distribuição</span>' }
            default { '<span class="badge-status badge-warning">Desconhecido</span>' }
        }
        
        $scopeBadge = switch ($group.Escopo) {
            "DomainLocal" { '<span class="badge-status badge-warning">Local de Domínio</span>' }
            "Global" { '<span class="badge-status badge-success">Global</span>' }
            "Universal" { '<span class="badge-status badge-info">Universal</span>' }
            default { '<span class="badge-status badge-warning">Desconhecido</span>' }
        }
        
        $memberCountBadge = if ($group.ContadorMembros -eq 0) {
            '<span class="badge-status badge-danger">Vazio</span>'
        } else {
            '<span class="badge-status badge-success">' + $group.ContadorMembros + '</span>'
        }
        
        $bodyContent += @"
                    <tr>
                        <td>$($group.Nome)</td>
                        <td class="truncate" title="$($group.Descrição)">$($group.Descrição)</td>
                        <td>$categoryBadge</td>
                        <td>$scopeBadge</td>
                        <td>$memberCountBadge</td>
                        <td class="truncate" title="$($group.Membros)">$($group.Membros)</td>
                        <td class="truncate" title="$($group.MembroDe)">$($group.MembroDe)</td>
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

# Preparar dados para gráfico Top 10 grupos
$topGroupLabels = "["
$topGroupData = "["

foreach ($group in $groupsWithMostMembers) {
    $topGroupLabels += "'$($group.Name)',"
    $topGroupData += "$($group.MemberCount),"
}

$topGroupLabels = $topGroupLabels.TrimEnd(',') + "]"
$topGroupData = $topGroupData.TrimEnd(',') + "]"

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
        
        // Gráfico de contagem de membros
        const memberCtx = document.getElementById('memberCountChart').getContext('2d');
        const memberChart = new Chart(memberCtx, {
            type: 'bar',
            data: {
                labels: $topGroupLabels,
                datasets: [{
                    label: 'Número de Membros',
                    data: $topGroupData,
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
                },
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });
    });

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
                      -Title "Active Directory Report Tool - Membros de Grupos" `
                      -ActiveMenu "Membros de Grupos" `
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