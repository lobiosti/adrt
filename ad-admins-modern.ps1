<#
.SYNOPSIS
    Active Directory - Administradores de Domínio (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para listar administradores de domínio no Active Directory
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-admins.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"
$directoryPath = (Get-Item -Path ".").FullName
#$outputPath = "ad-reports\ad-admins\ad-admins-modern.html"
# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-admins\ad-admins-modern.html"

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
Write-Host "Coletando dados de administradores de domínio..."
try {
    $admins = Get-ADGroupMember -Identity "Domain Admins" | 
              Get-ADUser -Properties Name, SamAccountName, EmailAddress, Title, Department, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires, Modified

    # Coletar estatísticas para o dashboard
    $totalAdmins = $admins.Count
    $enterpriseAdmins = 0
    try {
        $enterpriseAdmins = (Get-ADGroupMember -Identity "Enterprise Admins" -ErrorAction SilentlyContinue).Count
    } catch {
        # Grupo Enterprise Admins pode não existir
    }
    $passwordNeverExpires = ($admins | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
}
catch {
    Write-Host "Erro ao coletar informações de Domain Admins: $_" -ForegroundColor Red
    Write-Host "Você pode não ter permissões suficientes."
    $admins = @()
    $totalAdmins = 0
    $enterpriseAdmins = 0
    $passwordNeverExpires = 0
}

# Preparar os dados para o relatório
$adminData = @()
foreach ($admin in $admins) {
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
    <h1>Active Directory - Administradores de Domínio</h1>
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
                    <p><strong>Total de Administradores de Domínio:</strong> $totalAdmins</p>
                    <p><strong>Administradores com senha que nunca expira:</strong> $passwordNeverExpires</p>
                </div>
"@

    if ($passwordNeverExpires -gt 0) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$passwordNeverExpires administradores com senhas que nunca expiram</span>
                </div>
"@
    }

    if ($totalAdmins -gt 3) {
        $bodyContent += @"
                <div class="warning-item warning-red">
                    <i class="fas fa-exclamation-circle"></i>
                    <span>Número elevado de administradores de domínio ($totalAdmins)</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-info-circle"></i>
                    <span>Nenhum administrador de domínio encontrado ou você não tem permissão para visualizar</span>
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
    <div class="col-md-6">
        <div class="card stat-card">
            <i class="fas fa-user-shield"></i>
            <h3>$totalAdmins</h3>
            <p>Administradores de Domínio</p>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card stat-card">
            <i class="fas fa-user-tie"></i>
            <h3>$enterpriseAdmins</h3>
            <p>Administradores Enterprise</p>
        </div>
    </div>
</div>

<!-- Gráfico -->
<div class="card mb-4">
    <div class="card-header">Distribuição de Permissões Administrativas</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="adminChart"></canvas>
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
                        <th>Última Modificação</th>
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
                        <td>$($admin.Nome)</td>
                        <td>$($admin.Login)</td>
                        <td>$($admin.Email)</td>
                        <td>$($admin.Cargo)</td>
                        <td>$($admin.Departamento)</td>
                        <td>$statusBadge</td>
                        <td>$($admin.UltimoLogin)</td>
                        <td>$($admin.UltimaModificacao)</td>
                        <td>$passwordBadge</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewUser('$($admin.Login)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="editUser('$($admin.Login)')"><i class="fas fa-edit"></i></button>
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
            <i class="fas fa-info-circle"></i> Nenhum administrador de domínio encontrado ou você não tem permissão para visualizar.
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
    // Inicializar o gráfico
    document.addEventListener('DOMContentLoaded', function() {
        const ctx = document.getElementById('adminChart').getContext('2d');
        const adminChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Domain Admins', 'Enterprise Admins'],
                datasets: [{
                    label: 'Quantidade',
                    data: [$totalAdmins, $enterpriseAdmins],
                    backgroundColor: [
                        '#6a3094', // Roxo para Domain Admins
                        '#9657c7'  // Roxo claro para Enterprise Admins
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
            document.querySelector('.badge.badge-primary').textContent = visibleCount + ' registros';
        });
    });

    // Funções para interação com administradores
    function viewUser(username) {
        alert('Visualizando detalhes do administrador: ' + username);
        // Aqui poderia implementar a visualização real
    }
    
    function editUser(username) {
        alert('Editando administrador: ' + username);
        // Aqui poderia implementar a edição real
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
                     -Title "Active Directory Report Tool - Administradores de Domínio" `
                     -ActiveMenu "Administradores" `
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