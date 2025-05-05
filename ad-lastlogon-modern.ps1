<#
.SYNOPSIS
    Active Directory - Usuários sem Login Recente (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para Active Directory - Lista de usuários que não fizeram login nos últimos 90 dias
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-lastlogon.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-lastlogon\ad-lastlogon-modern.html"

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
Write-Host "Coletando dados de usuários sem login recente..."
try {
    # Parâmetros da análise de último login
    $days = 90
    $timestamp = (Get-Date).Adddays(-($days))
    
    # Obter usuários sem login recente (últimos 90 dias)
    $lastlogon = Get-ADUser -Filter {LastLogonTimeStamp -lt $timestamp -and enabled -eq $true} -Properties Name, SamAccountName, Title, Department, PasswordLastSet, LastLogonDate, LastLogonTimestamp, mail, Description, whenCreated
    
    # Coletar estatísticas para o dashboard
    $totalUsers = (Get-ADUser -Filter * -Properties Enabled | Where-Object { $_.Enabled -eq $true }).Count
    $inactiveUsers = $lastlogon.Count
    $inactivePercentage = if ($totalUsers -gt 0) { [math]::Round(($inactiveUsers / $totalUsers) * 100, 1) } else { 0 }
    
    # Coletar dados de usuários desativados para comparação
    $disabledUsers = (Get-ADUser -Filter { Enabled -eq $false }).Count
}
catch {
    Write-Host "Erro ao coletar informações de usuários sem login recente: $_" -ForegroundColor Red
    $lastlogon = @()
    $totalUsers = 0
    $inactiveUsers = 0
    $inactivePercentage = 0
    $disabledUsers = 0
}

# Preparar os dados para o relatório
$userData = @()
foreach ($user in $lastlogon) {
    # Converter timestamp para data legível
    $lastLogonDate = if ($user.LastLogonTimestamp) {
        [DateTime]::FromFileTime($user.LastLogonTimestamp).ToString('yyyy/MM/dd HH:mm:ss')
    } else {
        "Nunca"
    }
    
    # Calcular dias desde o último login
    $daysSinceLogin = if ($user.LastLogonDate) {
        [math]::Round(((Get-Date) - $user.LastLogonDate).TotalDays, 0)
    } else {
        "N/A"
    }
    
    $userData += [PSCustomObject]@{
        Nome = $user.Name
        Login = $user.SamAccountName
        Cargo = $user.Title
        Departamento = $user.Department
        Email = $user.mail
        UltimoLogin = $lastLogonDate
        DiasSemLogin = $daysSinceLogin
        SenhaConfigurada = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        DataCriacao = if ($user.whenCreated) { $user.whenCreated.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        Observacoes = $user.Description
    }
}

# Ordenar resultados por dias sem login (decrescente)
$userData = $userData | Sort-Object -Property DiasSemLogin -Descending

# Contar registros
$totalRecords = $userData.Count

# Gerar conteúdo do corpo do relatório
$bodyContent = @"
<!-- Cabeçalho padrão -->
<div class="header">
    <h1>Active Directory - Usuários sem Login ($days dias)</h1>
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
            <div class="card-header">Resumo da Segurança</div>
            <div class="card-body">
"@

if ($inactiveUsers -gt 0) {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Usuários sem login nos últimos $days dias:</strong> $inactiveUsers de $totalUsers ($inactivePercentage%)</p>
                </div>
"@

    if ($inactivePercentage -gt 20) {
        $bodyContent += @"
                <div class="warning-item warning-red">
                    <i class="fas fa-exclamation-circle"></i>
                    <span>Alto percentual de contas inativas ($inactivePercentage%). Recomendável revisar e considerar desativar.</span>
                </div>
"@
    } elseif ($inactivePercentage -gt 10) {
        $bodyContent += @"
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>Percentual significativo de contas inativas ($inactivePercentage%). Recomendável revisar.</span>
                </div>
"@
    }
} else {
    $bodyContent += @"
                <div class="info-box">
                    <p><strong>Situação:</strong> Nenhum usuário sem login recente encontrado.</p>
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
            <i class="fas fa-users"></i>
            <h3>$totalUsers</h3>
            <p>Total de Usuários Ativos</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-clock text-warning"></i>
            <h3>$inactiveUsers</h3>
            <p>Usuários sem Login ($days dias)</p>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card stat-card">
            <i class="fas fa-user-times text-danger"></i>
            <h3>$disabledUsers</h3>
            <p>Usuários Desativados</p>
        </div>
    </div>
</div>

<!-- Gráfico -->
<div class="card mb-4">
    <div class="card-header">Distribuição de Atividade de Usuários</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="userActivityChart"></canvas>
        </div>
    </div>
</div>

<!-- Tabela de Dados -->
<div class="card">
    <div class="card-header">
        <div>Usuários sem Login nos Últimos $days Dias</div>
        <div>
            <span class="badge badge-primary">$totalRecords registros</span>
        </div>
    </div>
    <div class="card-body">
"@

if ($inactiveUsers -gt 0) {
    $bodyContent += @"
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Login</th>
                        <th>Departamento</th>
                        <th>Cargo</th>
                        <th>Último Login</th>
                        <th>Dias sem Login</th>
                        <th>Data de Criação</th>
                        <th>Observações</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Adicionar dados à tabela
    foreach ($user in $userData) {
        # Definir classes para destacar contas com mais dias sem login
        $riskClass = ""
        if ($user.DiasSemLogin -ne "N/A") {
            if ([int]$user.DiasSemLogin -gt 180) {
                $riskClass = "class='risk-high'"
            } elseif ([int]$user.DiasSemLogin -gt 120) {
                $riskClass = "class='risk-medium'"
            }
        }
        
        $bodyContent += @"
                    <tr>
                        <td>$($user.Nome)</td>
                        <td>$($user.Login)</td>
                        <td>$($user.Departamento)</td>
                        <td>$($user.Cargo)</td>
                        <td>$($user.UltimoLogin)</td>
                        <td $riskClass>$($user.DiasSemLogin)</td>
                        <td>$($user.DataCriacao)</td>
                        <td>$($user.Observacoes)</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewUser('$($user.Login)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="disableUser('$($user.Login)')"><i class="fas fa-user-times"></i></button>
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
        <div class="alert alert-success">
            <i class="fas fa-check-circle"></i> Nenhum usuário encontrado sem login nos últimos $days dias.
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
        const ctx = document.getElementById('userActivityChart').getContext('2d');
        const activityChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Usuários Ativos', 'Sem Login ($days dias)', 'Desativados'],
                datasets: [{
                    data: [$totalUsers - $inactiveUsers, $inactiveUsers, $disabledUsers],
                    backgroundColor: [
                        '#28a745', // Verde para usuários ativos
                        '#ffc107', // Amarelo para sem login
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
    });

    // Funções para interação com usuários
    function viewUser(username) {
        alert('Visualizando detalhes do usuário: ' + username);
        // Aqui poderia redirecionar para uma página de detalhes
    }
    
    function disableUser(username) {
        if (confirm('Deseja realmente desativar o usuário ' + username + '?')) {
            alert('Usuário ' + username + ' seria desativado em um ambiente de produção.');
            // Aqui seria implementada a lógica para desativar o usuário
        }
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
                      -Title "Active Directory Report Tool - Usuários sem Login Recente" `
                      -ActiveMenu "Último Login" `
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