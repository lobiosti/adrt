<#
.SYNOPSIS
    Active Directory - Usuários Desativados (Simplificado)
.DESCRIPTION
    Script ADRT para listar usuários desativados no Active Directory
    Usando a abordagem de helper script em vez de módulo
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"
#$outputPath = "ad-reports\ad-disabled\ad-disabled-modern.html"


# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-disabled\ad-disabled-modern.html"

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
Write-Host "Coletando dados de usuários desativados..."
try {
    # Obter todos os usuários desativados
    $disabledUsers = Search-ADAccount -AccountDisabled -UsersOnly | 
                     Get-ADUser -Properties Name, SamAccountName, Created, Modified, LastLogonDate, UserPrincipalName, EmailAddress, Department, Title
    
    # Coletar estatísticas para o dashboard
    $totalDisabledUsers = $disabledUsers.Count
    $totalUsers = (Get-ADUser -Filter *).Count
    $percentDisabled = if ($totalUsers -gt 0) { [math]::Round(($totalDisabledUsers / $totalUsers) * 100, 1) } else { 0 }
    
    # Verificar usuários desativados recentemente (últimos 30 dias)
    $recentlyDisabled = ($disabledUsers | Where-Object { $_.Modified -ge (Get-Date).AddDays(-30) }).Count
    
    # Verificar usuários que nunca fizeram login
    $neverLoggedIn = ($disabledUsers | Where-Object { $_.LastLogonDate -eq $null }).Count
    
    # Verificar usuários desativados há mais de 6 meses
    $olderThan6Months = ($disabledUsers | Where-Object { $_.Modified -lt (Get-Date).AddMonths(-6) }).Count
}
catch {
    Write-Host "Erro ao coletar informações de usuários desativados: $_" -ForegroundColor Red
    $disabledUsers = @()
    $totalDisabledUsers = 0
    $totalUsers = 0
    $percentDisabled = 0
    $recentlyDisabled = 0
    $neverLoggedIn = 0
    $olderThan6Months = 0
}

# Preparar os dados para o relatório
$userData = @()
foreach ($user in $disabledUsers) {
    $userData += [PSCustomObject]@{
        Nome = $user.Name
        Login = $user.SamAccountName
        Email = if ($user.EmailAddress) { $user.EmailAddress } else { $user.UserPrincipalName }
        Departamento = $user.Department
        Cargo = $user.Title
        UltimoLogin = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        Criado = if ($user.Created) { $user.Created.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        DesativadoEm = if ($user.Modified) { $user.Modified.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        StatusDesativacao = if ($user.Modified -ge (Get-Date).AddDays(-30)) { "Recente" } 
                          elseif ($user.Modified -lt (Get-Date).AddMonths(-6)) { "Antigo" } 
                          else { "Normal" }
    }
}

# Ordenar resultados
$userData = $userData | Sort-Object -Property "Nome"

# Contar registros
$totalRecords = $userData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Usuários Desativados</h1>
    <div class="header-actions">
        <button onclick="filterUsers('all')"><i class="fas fa-sync"></i> Mostrar Todos</button>
        <button onclick="filterUsers('recent')"><i class="fas fa-history"></i> Desativados Recentemente</button>
        <button onclick="filterUsers('old')"><i class="fas fa-archive"></i> Desativados Há Muito Tempo</button>
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
            <div class="card-header">Resumo de Usuários Desativados</div>
            <div class="card-body">
"@

if ($totalDisabledUsers -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Usuários Desativados:</strong> $totalDisabledUsers de $totalUsers ($percentDisabled%)</p>
                    <p><strong>Desativados nos últimos 30 dias:</strong> $recentlyDisabled</p>
                    <p><strong>Desativados há mais de 6 meses:</strong> $olderThan6Months</p>
                    <p><strong>Nunca fizeram login:</strong> $neverLoggedIn</p>
                </div>
"@

    if ($percentDisabled -gt 20) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>Alta porcentagem de usuários desativados ($percentDisabled%)</span>
                </div>
"@
    }
    
    if ($olderThan6Months -gt ($totalDisabledUsers * 0.7)) {
        $oldPercent = [math]::Round(($olderThan6Months / $totalDisabledUsers) * 100, 1)
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$oldPercent% das contas desativadas estão inutilizadas há mais de 6 meses e poderiam ser removidas</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum usuário desativado encontrado ou você não tem permissão para visualizar</span>
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
            <i class="fas fa-user-times"></i>
            <h3>$totalDisabledUsers</h3>
            <p>Total de Usuários Desativados</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-history"></i>
            <h3>$recentlyDisabled</h3>
            <p>Desativados Recentemente</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-archive"></i>
            <h3>$olderThan6Months</h3>
            <p>Desativados Há Mais de 6 Meses</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Proporção de Contas</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="accountChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Tempo de Desativação</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="disabledTimeChart"></canvas>
                </div>
            </div>
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

if ($totalDisabledUsers -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="disabledUsersTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Login</th>
                        <th>Email</th>
                        <th>Departamento</th>
                        <th>Cargo</th>
                        <th>Último Login</th>
                        <th>Desativado Em</th>
                        <th>Status</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($user in $userData) {
        $statusBadge = switch ($user.StatusDesativacao) {
            "Recente" { '<span class="badge-status badge-recent" data-status="recent">Recente</span>' }
            "Antigo" { '<span class="badge-status badge-old" data-status="old">Antigo</span>' }
            default { '<span class="badge-status badge-warning">Normal</span>' }
        }
        
        $bodyContent += @"
                    <tr data-status="$($user.StatusDesativacao.ToLower())">
                        <td>$($user.Nome)</td>
                        <td>$($user.Login)</td>
                        <td>$($user.Email)</td>
                        <td>$($user.Departamento)</td>
                        <td>$($user.Cargo)</td>
                        <td>$($user.UltimoLogin)</td>
                        <td>$($user.DesativadoEm)</td>
                        <td>$statusBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewUser('$($user.Login)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="enableUser('$($user.Login)')"><i class="fas fa-user-check"></i></button>
                            <button class="action-button" onclick="removeUser('$($user.Login)')"><i class="fas fa-trash-alt"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhum usuário desativado encontrado ou você não tem permissão para visualizar.
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
        // Gráfico de proporção de contas
        const accountCtx = document.getElementById('accountChart').getContext('2d');
        const accountChart = new Chart(accountCtx, {
            type: 'pie',
            data: {
                labels: ['Usuários Ativos', 'Usuários Desativados'],
                datasets: [{
                    data: [$totalUsers - $totalDisabledUsers, $totalDisabledUsers],
                    backgroundColor: [
                        '#28a745', // Verde para Ativos
                        '#dc3545'  // Vermelho para Desativados
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
        
        // Gráfico de tempo de desativação
        const timeCtx = document.getElementById('disabledTimeChart').getContext('2d');
        const timeChart = new Chart(timeCtx, {
            type: 'pie',
            data: {
                labels: ['Recentes (< 30 dias)', 'Normais (30d - 6m)', 'Antigos (> 6 meses)'],
                datasets: [{
                    data: [$recentlyDisabled, $totalDisabledUsers - $recentlyDisabled - $olderThan6Months, $olderThan6Months],
                    backgroundColor: [
                        '#d63384',  // Rosa para Recentes
                        '#fd7e14',  // Laranja para Normais
                        '#6c757d'   // Cinza para Antigos
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
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#disabledUsersTable tbody tr');
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

    // Função para filtrar usuários por status de desativação
    function filterUsers(status) {
        const rows = document.querySelectorAll('#disabledUsersTable tbody tr');
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
    
    // Funções para interação com usuários
    function viewUser(username) {
        alert('Visualizando detalhes do usuário: ' + username);
        // Aqui poderia implementar a visualização real
    }
    
    function enableUser(username) {
        if (confirm('Deseja reativar o usuário ' + username + '?')) {
            alert('Reativando usuário: ' + username);
            // Aqui poderia implementar a reativação real
        }
    }
    
    function removeUser(username) {
        if (confirm('Deseja remover permanentemente o usuário ' + username + '?')) {
            alert('Removendo usuário: ' + username);
            // Aqui poderia implementar a remoção real do usuário
        }
    }
</script>
"@

# Parâmetros para o template
$templateParams = @{
    Title = "Active Directory Report Tool - Usuários Desativados"
    ActiveMenu = "Usuários Desativados"
    CompanyName = $company
    DomainName = (Get-ADDomain).Forest
    Date = $date
    Owner = $owner
    ExtraScripts = $extraScripts
}

# Gerar HTML utilizando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Usuários Desativados" `
                      -ActiveMenu "Usuários Desativados" `
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