<#
.SYNOPSIS
    Active Directory - Administradores Enterprise (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para Active Directory - Lista de administradores enterprise
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-enterprise-admins.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-enterprise-admins\ad-enterprise-admins-modern.html"

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
Write-Host "Coletando dados de administradores enterprise..."
try {
    # Obter membros do grupo Enterprise Admins
    $admins = Get-ADGroupMember -Identity "Enterprise Admins" -ErrorAction SilentlyContinue | 
              Get-ADUser -Properties Name, SamAccountName, EmailAddress, Title, Department, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires, Modified, Description, whenCreated

    # Coletar estatísticas para o dashboard
    $totalAdmins = $admins.Count
    $domainAdmins = (Get-ADGroupMember -Identity "Domain Admins" -ErrorAction SilentlyContinue).Count
    $passwordNeverExpires = ($admins | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
    $disabledAdmins = ($admins | Where-Object { $_.Enabled -eq $false }).Count
    $activeAdmins = $totalAdmins - $disabledAdmins
    
    # Calcular percentual de administradores com senha que nunca expira
    $percentNeverExpires = if ($totalAdmins -gt 0) {
        [math]::Round(($passwordNeverExpires / $totalAdmins) * 100, 1)
    } else {
        0
    }
    
    # Verificar administradores sem login recente (90 dias)
    $daysThreshold = 90
    $cutoffDate = (Get-Date).AddDays(-$daysThreshold)
    $inactiveAdmins = ($admins | Where-Object { 
        ($_.LastLogonDate -lt $cutoffDate -or $_.LastLogonDate -eq $null) -and $_.Enabled -eq $true 
    }).Count
    
    # Verificar estatísticas de segurança
    $totalAccounts = (Get-ADUser -Filter *).Count
    $adminRatio = if ($totalAccounts -gt 0) {
        [math]::Round(($totalAdmins / $totalAccounts) * 100, 2)
    } else {
        0
    }
}
catch {
    Write-Host "Erro ao coletar informações de Enterprise Admins: $_" -ForegroundColor Red
    Write-Host "O grupo Enterprise Admins pode não existir neste domínio ou você não tem permissões suficientes."
    $admins = @()
    $totalAdmins = 0
    $domainAdmins = 0 
    $passwordNeverExpires = 0
    $disabledAdmins = 0
    $activeAdmins = 0
    $percentNeverExpires = 0
    $inactiveAdmins = 0
    $adminRatio = 0
    $totalAccounts = 0
}

# Preparar os dados para o relatório
$adminData = @()
foreach ($admin in $admins) {
    # Calcular dias desde o último login
    $daysSinceLogin = if ($admin.LastLogonDate) {
        [math]::Round(((Get-Date) - $admin.LastLogonDate).TotalDays, 0)
    } else {
        "N/A"
    }
    
    # Determinar status de segurança
    $securityStatus = "Normal"
    if ($admin.PasswordNeverExpires -eq $true) {
        $securityStatus = "Atenção"
    }
    if ($daysSinceLogin -ne "N/A" -and $daysSinceLogin -gt $daysThreshold) {
        $securityStatus = "Atenção"
    }
    if (-not $admin.Enabled) {
        $securityStatus = "Desativado"
    }
    
    $adminData += [PSCustomObject]@{
        Nome = $admin.Name
        Login = $admin.SamAccountName
        Email = $admin.EmailAddress
        Cargo = $admin.Title
        Departamento = $admin.Department
        Status = if ($admin.Enabled) { "Ativo" } else { "Inativo" }
        UltimoLogin = if ($admin.LastLogonDate) { $admin.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        UltimaModificacao = if ($admin.Modified) { $admin.Modified.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        SenhaConfigurada = if ($admin.PasswordLastSet) { $admin.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        SenhaNuncaExpira = $admin.PasswordNeverExpires
        DiasSemLogin = $daysSinceLogin
        DataCriacao = if ($admin.whenCreated) { $admin.whenCreated.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Observacoes = $admin.Description
        StatusSeguranca = $securityStatus
    }
}

# Ordenar resultados
$adminData = $adminData | Sort-Object -Property "Nome"

# Contar registros
$totalRecords = $adminData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Administradores Enterprise</h1>
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

if ($totalAdmins -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Total de Administradores Enterprise:</strong> $totalAdmins</p>
                    <p><strong>Administradores Ativos:</strong> $activeAdmins</p>
                    <p><strong>Administradores com senha que nunca expira:</strong> $passwordNeverExpires ($percentNeverExpires%)</p>
                    <p><strong>Percentual de contas Enterprise Admins:</strong> $adminRatio% do total de contas</p>
                </div>
"@

    if ($passwordNeverExpires -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$passwordNeverExpires administradores com senhas que nunca expiram.</span>
                </div>
"@
    }

    if ($inactiveAdmins -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$inactiveAdmins administradores sem login nos últimos $daysThreshold dias.</span>
                </div>
"@
    }

    if ($totalAdmins -gt 2) {
        $bodyContent += @"
                <div class="warning-item warning-red">
                    <i class="fas fa-exclamation-circle"></i>
                    <span>Número elevado de administradores enterprise ($totalAdmins). Recomenda-se manter apenas 1-2 contas.</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum administrador enterprise encontrado neste domínio ou você não tem permissão para visualizar.</span>
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
            <i class="fas fa-user-tie"></i>
            <h3>$totalAdmins</h3>
            <p>Administradores Enterprise</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-user-shield"></i>
            <h3>$domainAdmins</h3>
            <p>Administradores de Domínio</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-key text-warning"></i>
            <h3>$passwordNeverExpires</h3>
            <p>Senhas Permanentes</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-user-lock text-danger"></i>
            <h3>$disabledAdmins</h3>
            <p>Desativados</p>
        </div>
    </div>
</div>

<!-- Gráficos -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribuição de Permissões Administrativas</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="adminDistributionChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Status de Segurança dos Administradores</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="adminSecurityChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Administradores Enterprise</div>
        <div>
            <span class="badge badge-primary" id="recordCount">$totalRecords registros</span>
            <input type="text" id="searchInput" placeholder="Filtrar..." class="form-control" style="display: inline-block; width: 200px; margin-left: 10px;">
        </div>
    </div>
    <div class="card-body">
"@

if ($totalAdmins -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table id="adminsTable">
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Login</th>
                        <th>Email</th>
                        <th>Cargo</th>
                        <th>Departamento</th>
                        <th>Status</th>
                        <th>Último Login</th>
                        <th>Senha Nunca Expira</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($admin in $adminData) {
        $statusBadge = if ($admin.Status -eq "Ativo") {
            '<span class="badge-status badge-success">Ativo</span>'
        } else {
            '<span class="badge-status badge-danger">Inativo</span>'
        }
        
        $passwordBadge = if ($admin.SenhaNuncaExpira) {
            '<span class="badge-status badge-danger">Nunca Expira</span>'
        } else {
            '<span class="badge-status badge-success">Normal</span>'
        }
        
        $bodyContent += @"
                    <tr>
                        <td><i class="fas fa-user-tie"></i> $($admin.Nome)</td>
                        <td>$($admin.Login)</td>
                        <td>$($admin.Email)</td>
                        <td>$($admin.Cargo)</td>
                        <td>$($admin.Departamento)</td>
                        <td>$statusBadge</td>
                        <td>$($admin.UltimoLogin)</td>
                        <td>$passwordBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewUser('$($admin.Login)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="editUser('$($admin.Login)')"><i class="fas fa-edit"></i></button>
                            <button class="action-button" onclick="resetPassword('$($admin.Login)')"><i class="fas fa-key"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhum administrador enterprise encontrado neste domínio ou você não tem permissão para visualizar.
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
        // Gráfico de distribuição de administradores
        const distributionCtx = document.getElementById('adminDistributionChart').getContext('2d');
        const distributionChart = new Chart(distributionCtx, {
            type: 'bar',
            data: {
                labels: ['Enterprise Admins', 'Domain Admins'],
                datasets: [{
                    label: 'Quantidade de Administradores',
                    data: [$totalAdmins, $domainAdmins],
                    backgroundColor: [
                        '#6a3094', // Roxo para Enterprise Admins
                        '#9657c7'  // Roxo claro para Domain Admins
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
        
        // Gráfico de segurança de administradores
        const securityCtx = document.getElementById('adminSecurityChart').getContext('2d');
        const securityChart = new Chart(securityCtx, {
            type: 'pie',
            data: {
                labels: ['Administradores Ativos', 'Administradores Desativados', 'Senhas Nunca Expiram', 'Sem Login ($daysThreshold dias)'],
                datasets: [{
                    data: [
                        $($activeAdmins - $passwordNeverExpires - $inactiveAdmins),
                        $disabledAdmins,
                        $passwordNeverExpires,
                        $inactiveAdmins
                    ],
                    backgroundColor: [
                        '#28a745', // Verde para ativos sem problemas
                        '#dc3545', // Vermelho para desativados
                        '#ffc107', // Amarelo para senhas que nunca expiram
                        '#fd7e14'  // Laranja para inativos
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
            const rows = document.querySelectorAll('#adminsTable tbody tr');
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

    // Funções para interação com usuários
    function viewUser(username) {
        alert('Visualizando detalhes do administrador: ' + username);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function editUser(username) {
        alert('Editando administrador: ' + username);
        // Aqui poderia abrir um formulário de edição
    }
    
    function resetPassword(username) {
        if (confirm('Deseja realmente redefinir a senha do administrador ' + username + '?')) {
            alert('Senha do administrador ' + username + ' seria redefinida em um ambiente de produção.');
            // Aqui seria implementada a lógica para redefinir a senha
        }
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
                      -Title "Active Directory Report Tool - Administradores Enterprise" `
                      -ActiveMenu "Administradores Enterprise" `
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