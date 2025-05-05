<#
.SYNOPSIS
    Active Directory - Senhas que Nunca Expiram (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para Active Directory - Lista de usuários com senhas que nunca expiram
    Utilizando o ADRT-Helper.ps1 para geração do relatório
.NOTES
    Original: ad-neverexpires.ps1
    Convertido para formato moderno e otimizado
#>

# Variáveis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obtém o diretório onde o script está localizado, não o diretório atual de execução
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-neverexpires\ad-neverexpires-modern.html"

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
Write-Host "Coletando dados de usuários com senhas que nunca expiram..."
try {
    # Obter usuários com senha que nunca expira
    $neverExpires = Get-ADUser -Filter * -Properties Name, SamAccountName, PasswordNeverExpires, Enabled, Title, Department, LastLogonDate, PasswordLastSet, Mail | 
                    Where-Object { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true }
    
    # Coletar estatísticas para o dashboard
    $totalUsers = (Get-ADUser -Filter *).Count
    $disabledUsers = (Search-ADAccount -AccountDisabled -UsersOnly).Count
    $activeUsers = $totalUsers - $disabledUsers
    $usersWithNeverExpires = $neverExpires.Count
    
    # Coletar usuários sem login recente (corrigido)
    $timestamp = (Get-Date).AddDays(-90)
    $lastLogon90Days = (Get-ADUser -Filter { Enabled -eq $true } -Properties LastLogonDate | 
                        Where-Object { $_.LastLogonDate -lt $timestamp -or $_.LastLogonDate -eq $null }).Count
}
catch {
    Write-Host "Erro ao coletar dados de usuários com senhas que nunca expiram: $_" -ForegroundColor Red
    $neverExpires = @()
    $totalUsers = 0
    $disabledUsers = 0
    $activeUsers = 0
    $usersWithNeverExpires = 0
    $lastLogon90Days = 0
}

# Preparar os dados para o relatório
$userData = @()
foreach ($user in $neverExpires) {
    $userData += [PSCustomObject]@{
        Nome = $user.Name
        Login = $user.SamAccountName
        Email = $user.Mail
        Cargo = $user.Title
        Departamento = $user.Department
        UltimoLogin = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" }
        SenhaConfigurada = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Desconhecido" }
        NuncaExpira = $user.PasswordNeverExpires
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
    <h1>Active Directory - Senhas que Nunca Expiram</h1>
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
                <div class="warning-item warning-yellow">
                    <i class="fas fa-exclamation-triangle"></i>
                    <span>$usersWithNeverExpires usuários com senhas que nunca expiram</span>
                </div>
                <div class="warning-item warning-red">
                    <i class="fas fa-exclamation-circle"></i>
                    <span>$lastLogon90Days usuários não fizeram login nos últimos 90 dias</span>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Gráfico -->
<div class="card mb-4">
    <div class="card-header">Status de Segurança</div>
    <div class="card-body">
        <div class="chart-container">
            <canvas id="securityChart"></canvas>
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
        <div class="table-responsive">
            <table>
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Login</th>
                        <th>Email</th>
                        <th>Cargo</th>
                        <th>Departamento</th>
                        <th>Último Login</th>
                        <th>Senha Configurada</th>
                        <th>Expiração de Senha</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
"@

# Adicionar dados à tabela
foreach ($user in $userData) {
    $expirationStatus = if ($user.NuncaExpira) {
        '<span class="badge-status badge-danger">Nunca Expira</span>'
    } else {
        '<span class="badge-status badge-success">Expira Normalmente</span>'
    }
    
    $bodyContent += @"
                    <tr>
                        <td>$($user.Nome)</td>
                        <td>$($user.Login)</td>
                        <td>$($user.Email)</td>
                        <td>$($user.Cargo)</td>
                        <td>$($user.Departamento)</td>
                        <td>$($user.UltimoLogin)</td>
                        <td>$($user.SenhaConfigurada)</td>
                        <td>$expirationStatus</td>
                        <td class="action-buttons">
                            <button class="action-button" onclick="viewUser('$($user.Login)')"><i class="fas fa-eye"></i></button>
                            <button class="action-button" onclick="editUser('$($user.Login)')"><i class="fas fa-edit"></i></button>
                        </td>
                    </tr>
"@
}

# Fechar o HTML
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
        const ctx = document.getElementById('securityChart').getContext('2d');
        const securityChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Senhas Nunca Expiram', 'Usuários Ativos', 'Usuários Desativados', 'Último Login > 90 dias'],
                datasets: [{
                    label: 'Quantidade',
                    data: [$usersWithNeverExpires, $activeUsers, $disabledUsers, $lastLogon90Days],
                    backgroundColor: [
                        '#28a745', // Verde
                        '#dc3545', // Vermelho
                        '#ffc107', // Amarelo
                        '#fd7e14'  // Laranja
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
    
    // Função para exportar para CSV
    function exportToCsv() {
        alert('Exportando dados para CSV...');
        // Implementação da exportação CSV
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - Senhas que Nunca Expiram" `
                      -ActiveMenu "Senhas Nunca Expiram" `
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