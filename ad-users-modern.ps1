<#
.SYNOPSIS
    Active Directory - Todos os Usuários (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar usuários no Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-users.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"
$directoryPath = (Get-Item -Path ".").FullName

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-users\ad-users-modern.html"

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
Write-Host "Coletando dados de usuários do Active Directory..."
try {
    # Obter todos os usuários - apenas com propriedades comuns para evitar erros
    $users = Get-ADUser -Filter * -Properties Company, SamAccountName, Name, Mail, Department, Title, PasswordNeverExpires, Enabled, Created, Modified, Info, LastLogonDate, whenCreated, Manager, PasswordLastSet
 
    # Contar totais
    $totalUsers = $users.Count
    $enabledUsers = ($users | Where-Object { $_.Enabled -eq $true }).Count
    $disabledUsers = $totalUsers - $enabledUsers
    $passwordNeverExpires = ($users | Where-Object { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true }).Count 

    # Contar usuários sem login recente (90 dias)
    $daysThreshold = 90
    $inactiveDate = (Get-Date).AddDays(-$daysThreshold)
    $inactiveUsers = ($users | Where-Object { $_.LastLogonDate -lt $inactiveDate -and $_.Enabled -eq $true }).Count   

    # Estatísticas de departamentos
    $departments = @{}
    foreach ($user in $users) {
        if ($user.Department) {
            $dept = $user.Department
            if ($departments.ContainsKey($dept)) {
                $departments[$dept]++
            } else {
                $departments[$dept] = 1
            }
        }
    }  

    # Top 5 departamentos
    $topDepartments = $departments.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5
}
catch {
    Write-Host "Erro ao coletar informações de usuários: $_" -ForegroundColor Red
    $users = @()
    $totalUsers = 0
    $enabledUsers = 0
    $disabledUsers = 0
    $passwordNeverExpires = 0
    $inactiveUsers = 0
    $departments = @{}
    $topDepartments = @()
}

# Preparar os dados para o relatório
$userData = @()
foreach ($user in $users) {
    # Calcular dias desde o último login
    $daysSinceLogin = if ($user.LastLogonDate) {
        [math]::Round(((Get-Date) - $user.LastLogonDate).TotalDays, 0)
    } else {
        "N/A"
    }
  
    # Determinar status de segurança
    $securityStatus = "Normal"
    if ($user.PasswordNeverExpires -eq $true -and $user.Enabled -eq $true) {
        $securityStatus = "Atenção"
    }
    if ($daysSinceLogin -ne "N/A" -and $daysSinceLogin -gt $daysThreshold -and $user.Enabled -eq $true) {
        $securityStatus = "Atenção"
    }
    if (-not $user.Enabled) {
        $securityStatus = "Desativado"
    }  

    $userData += [PSCustomObject]@{
        Nome = $user.Name
        Login = $user.SamAccountName
        Email = $user.Mail
        Departamento = $user.Department
        Cargo = $user.Title
        Empresa = $user.Company
        Status = $user.Enabled ? "Ativo" : "Desativado"
        UltimoLogin = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        DiasSemLogin = $daysSinceLogin
        SenhaNuncaExpira = $user.PasswordNeverExpires
        UltimaTrocaSenha = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Nunca" }
        DataCriacao = if ($user.whenCreated) { $user.whenCreated.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Observacoes = $user.Info
        StatusSeguranca = $securityStatus
    }
}

# Ordenar resultados por nome
$userData = $userData | Sort-Object -Property Nome

# Contar registros
$totalRecords = $userData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Todos os Usuários</h1>
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
            <div class="card-header">Resumo da Segurança</div>
            <div class="card-body">
"@

if ($totalUsers -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Usuários:</strong> $totalUsers</p>
                    <p><strong>Usuários Ativos:</strong> $enabledUsers</p>
                    <p><strong>Usuários Desativados:</strong> $disabledUsers</p>
                </div>
"@

    if ($passwordNeverExpires -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$passwordNeverExpires usuários com senhas que nunca expiram.</span>
                </div>
"@
    }
    
    if ($inactiveUsers -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$inactiveUsers usuários sem login nos últimos $daysThreshold dias.</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum usuário encontrado ou você não tem permissão para visualizar.</span>
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
            <h3>$totalUsers</h3>
            <p>Total de Usuários</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-user-check text-success"></i>
            <h3>$enabledUsers</h3>
            <p>Usuários Ativos</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-user-times text-danger"></i>
            <h3>$disabledUsers</h3>
            <p>Usuários Desativados</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-key text-warning"></i>
            <h3>$passwordNeverExpires</h3>
            <p>Senhas Permanentes</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
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
            <div class="card-header">Distribuição por Departamento</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="departmentChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Lista de Usuários</div>
        <div>
            <span class="badge badge-primary">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
"@

if ($totalUsers -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="usersTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Login</th>
                        <th>Email</th>
                        <th>Departamento</th>
                        <th>Cargo</th>
                        <th>Último Login</th>
                        <th>Status</th>
                        <th>Senha Nunca Expira</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($user in $userData) {
        $statusBadge = ""
        if ($user.Status -eq "Ativo") {
            $statusBadge = '<span class="badge-status badge-success">Ativo</span>'
        } else {
            $statusBadge = '<span class="badge-status badge-danger">Desativado</span>'
        }
        
        $passwordBadge = ""
        if ($user.SenhaNuncaExpira) {
            $passwordBadge = '<span class="badge-status badge-warning">Sim</span>'
        } else {
            $passwordBadge = '<span class="badge-status badge-success">Não</span>'
        }
        
        $bodyContent += @"
                    <tr>
                        <td><i class="fas fa-user"></i> $($user.Nome)</td>
                        <td>$($user.Login)</td>
                        <td>$($user.Email)</td>
                        <td>$($user.Departamento)</td>
                        <td>$($user.Cargo)</td>
                        <td>$($user.UltimoLogin)</td>
                        <td>$statusBadge</td>
                        <td>$passwordBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewUser('$($user.Login)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="editUser('$($user.Login)')"><i class="fas fa-edit"></i></button>
                            <button class="action-button" onclick="resetPassword('$($user.Login)')"><i class="fas fa-key"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhum usuário encontrado neste domínio ou você não tem permissão para visualizar.
        </div>
"@
}

$bodyContent += @"
    </div>
</div>
"@

# Preparar dados para os gráficos
$deptLabels = "["
$deptData = "["

foreach ($dept in $topDepartments) {
    $deptName = $dept.Key
    $deptLabels += "'$deptName',"
    $deptData += "$($dept.Value)," 
}

# Remover vírgulas finais e fechar arrays
$deptLabels = $deptLabels.TrimEnd(',') + "]"
$deptData = $deptData.TrimEnd(',') + "]"

# Script específico para esta página
$extraScripts = @"
<script>
    // Inicializar os gráficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gráfico de status de usuários
        const statusCtx = document.getElementById('userStatusChart').getContext('2d');
        const statusChart = new Chart(statusCtx, {
            type: 'pie',
            data: {
                labels: ['Ativos', 'Desativados'],
                datasets: [{
                    data: [$enabledUsers, $disabledUsers],
                    backgroundColor: [
                        '#28a745', // Verde para ativos
                        '#dc3545'  // Vermelho para desativados
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
        
        // Gráfico de departamentos
        const deptCtx = document.getElementById('departmentChart').getContext('2d');
        const deptChart = new Chart(deptCtx, {
            type: 'bar',
            data: {
                labels: $deptLabels,
                datasets: [{
                    label: 'Usuários por Departamento',
                    data: $deptData,
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
        
        // Adicionar funcionalidade de pesquisa
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#usersTable tbody tr');
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

    // Funções para interação com usuários
    function viewUser(username) {
        alert('Visualizando detalhes do usuário: ' + username);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function editUser(username) {
        alert('Editando usuário: ' + username);
        // Aqui poderia abrir um formulário de edição
    }
    
    function resetPassword(username) {
        if (confirm('Deseja realmente redefinir a senha do usuário ' + username + '?')) {
            alert('Senha do usuário ' + username + ' seria redefinida em um ambiente de produção.');
            // Aqui seria implementada a lógica para redefinir a senha
        }
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                       -Title "Active Directory Report Tool - Usuários" `
                       -ActiveMenu "Usuários" `
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