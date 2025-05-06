<#
.SYNOPSIS
    Active Directory - An�lise Completa (Formato Otimizado)
.DESCRIPTION
    Script ADRT modernizado para an�lise completa do Active Directory
    Utilizando o ADRT-Helper.ps1 para gera��o do relat�rio detalhado
.NOTES
    Original: analise-completa.ps1
    Convertido para formato moderno e otimizado
#>

# Definir codifica��o para garantir acentua��o correta
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding'] = 'UTF8'

# Vari�veis do script
$date = Get-Date -Format "yyyy-MM-dd"

# Obt�m o diret�rio onde o script est� localizado, n�o o diret�rio atual de execu��o
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$directoryPath = $scriptDir
$outputPath = Join-Path -Path $scriptDir -ChildPath "ad-reports\ad-analysis\ad-analysis-modern.html"

# Criar diret�rio se n�o existir
$outputDir = Split-Path -Path $outputPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "? Diret�rio de sa�da criado: $outputDir" -ForegroundColor Green
}

# Banner
Write-Host @"

?????????????????????????????????????????????????????????????????
?                                                               ?
?      ???      ??????? ??????? ??? ??????? ????????           ?
?      ???     ?????????????????????????????????????           ?
?      ???     ???   ?????????????????   ???????????           ?
?      ???     ???   ?????????????????   ???????????           ?
?      ?????????????????????????????????????????????           ?
?      ???????? ??????? ??????? ??? ??????? ????????           ?
?                                                               ?
?        ADRT - An�lise Completa do Active Directory            ?
?                                                               ?
?????????????????????????????????????????????????????????????????
"@ -ForegroundColor Magenta

# Obter informa��es de configura��o
if (Test-Path -Path "config\config.txt") {
    try {
        $config = Get-Content -Path "config\config.txt" -Encoding UTF8 -ErrorAction Stop
        $company = $config[7]
        $owner = $config[9]
        Write-Host "? Arquivo de configura��o carregado com sucesso" -ForegroundColor Green
    }
    catch {
        Write-Host "! Erro ao ler arquivo de configura��o. Usando valores padr�o." -ForegroundColor Yellow
        $company = "Lobios"
        $owner = "Administrador"
    }
}
else {
    Write-Host "! Arquivo de configura��o n�o encontrado. Usando valores padr�o." -ForegroundColor Yellow
    $company = "Lobios"
    $owner = "Administrador"
}

# Adicione este trecho antes do resumo final no atualizar-relatorios.ps1

# Tentar importar o m�dulo de notifica��o
$notificationModuleAvailable = $false
try {
    Import-Module ".\modules\ADRT-Notification.psm1" -ErrorAction Stop
    $notificationModuleAvailable = $true
    Write-Host ""
    Write-Host "Enviando notifica��es..." -ForegroundColor Cyan
}
catch {
    Write-Host "! M�dulo de notifica��es n�o encontrado. As notifica��es n�o ser�o enviadas." -ForegroundColor Yellow
}

# Enviar notifica��es se o m�dulo estiver dispon�vel
if ($notificationModuleAvailable) {
    try {
        # Criar hashtable para simular as estat�sticas necess�rias
        $stats = @{
            TotalUsers = 0
            TotalComputers = 0
            TotalServers = 0
            TotalGroups = 0
            TotalOUs = 0
            TotalGPOs = 0
            TotalDevices = 0 # Adicionado para evitar o erro
            DomainName = (Get-ADDomain).Forest
        }
        
        # Adicionar estat�sticas espec�ficas de atualiza��o
        $stats.AtualizacaoTotal = $totalScripts
        $stats.AtualizacaoSucessos = $sucessos
        $stats.AtualizacaoFalhas = $falhas
        
        # Caminho do index-modern.html
        $indexPath = Join-Path -Path (Get-Location).Path -ChildPath "index-modern.html"
        
        # Enviar notifica��o usando a fun��o existente
        $notificationSent = Send-ADRTNotification -ScriptName "atualizar-relatorios.ps1" `
                                                 -Type "Atualiza��o de Relat�rios" `
                                                 -Stats $stats `
                                                 -Domain $stats.DomainName `
                                                 -ReportPath $indexPath
        
        Write-Host "? Notifica��es enviadas com sucesso" -ForegroundColor Green
    }
    catch {
        Write-Host "? Erro ao enviar notifica��es: $_" -ForegroundColor Red
        $notificationSent = $false
    }
}

# Carregar o helper
. ".\modules\ADRT-Helper.ps1"

# Importar m�dulo ActiveDirectory
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "? M�dulo ActiveDirectory carregado com sucesso" -ForegroundColor Green
}
catch {
    Write-Host "? Erro cr�tico: N�o foi poss�vel carregar o m�dulo ActiveDirectory" -ForegroundColor Red
    Write-Host "Este script requer o m�dulo ActiveDirectory. Verifique se as ferramentas RSAT est�o instaladas." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "Iniciando an�lise completa do Active Directory..." -ForegroundColor Cyan
Write-Host "Coletando estat�sticas e m�tricas..." -ForegroundColor Cyan

# Coletar dados principais - inicializar estrutura
$stats = @{
    TotalUsers = 0
    EnabledUsers = 0
    DisabledUsers = 0
    Days = 90
    LastLogon90Days = 0
    PasswordNeverExpires = 0
    TotalComputers = 0
    TotalServers = 0
    TotalGroups = 0
    TotalOUs = 0
    DomainAdmins = 0
    EnterpriseAdmins = 0
    DomainControllers = 0
    TotalGPOs = 0
    DomainName = ""
    ForestLevel = ""
    DomainLevel = ""
}

# Contagens b�sicas com tratamento de erro
try {
    $stats.TotalUsers = (Get-ADUser -Filter *).Count
    $stats.EnabledUsers = (Get-ADUser -Filter {Enabled -eq $true}).Count
    $stats.DisabledUsers = $stats.TotalUsers - $stats.EnabledUsers
    Write-Host "Total de usu�rios: $($stats.TotalUsers)" -ForegroundColor Green
    Write-Host "Usu�rios ativos: $($stats.EnabledUsers)" -ForegroundColor Green
    Write-Host "Usu�rios desativados: $($stats.DisabledUsers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar usu�rios: $_" -ForegroundColor Yellow
}

# Calcular usu�rios com senha nunca expira
try {
    $stats.PasswordNeverExpires = (Get-ADUser -filter * -properties PasswordNeverExpires | 
        Where-Object { $_.PasswordNeverExpires -eq "true" -and $_.enabled -eq "true" }).Count
    Write-Host "Usu�rios com senha que nunca expira: $($stats.PasswordNeverExpires)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar usu�rios com senha que nunca expira: $_" -ForegroundColor Yellow
}

# Calcular usu�rios sem login nos �ltimos 90 dias
try {
    $timestamp = (Get-Date).AddDays(-($stats.Days))
    $stats.LastLogon90Days = (Get-ADUser -Filter {LastLogonTimeStamp -lt $timestamp -and enabled -eq $true} -Properties LastLogonTimeStamp).Count
    Write-Host "Usu�rios sem login nos �ltimos 90 dias: $($stats.LastLogon90Days)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar usu�rios sem login recente: $_" -ForegroundColor Yellow
}

# Computadores e servidores
try {
    $stats.TotalComputers = (Get-ADComputer -Filter { OperatingSystem -NotLike '*Windows Server*' }).Count
    Write-Host "Total de computadores: $($stats.TotalComputers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar computadores: $_" -ForegroundColor Yellow
}

try {
    $stats.TotalServers = (Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' }).Count
    Write-Host "Total de servidores: $($stats.TotalServers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar servidores: $_" -ForegroundColor Yellow
}

# Grupos e OUs
try {
    $stats.TotalGroups = (Get-ADGroup -Filter {name -like "*"}).Count
    Write-Host "Total de grupos: $($stats.TotalGroups)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar grupos: $_" -ForegroundColor Yellow
}

try {
    $stats.TotalOUs = (Get-ADOrganizationalUnit -Filter {name -like "*"}).Count
    Write-Host "Total de OUs: $($stats.TotalOUs)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar OUs: $_" -ForegroundColor Yellow
}

# Domain Controllers
try {
    $domainControllers = Get-ADDomainController -Filter * 
    $stats.DomainControllers = $domainControllers.Count
    Write-Host "Total de controladores de dom�nio: $($stats.DomainControllers)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar controladores de dom�nio: $_" -ForegroundColor Yellow
}

# Obter informa��es de dom�nio
try {
    $domain = Get-ADDomain
    $forest = Get-ADForest
    $stats.DomainName = $domain.DNSRoot
    $stats.DomainLevel = $domain.DomainMode
    $stats.ForestLevel = $forest.ForestMode
    Write-Host "Dom�nio: $($stats.DomainName)" -ForegroundColor Green
    Write-Host "N�vel funcional de dom�nio: $($stats.DomainLevel)" -ForegroundColor Green
    Write-Host "N�vel funcional de floresta: $($stats.ForestLevel)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao obter informa��es de dom�nio: $_" -ForegroundColor Yellow
}

# Domain Admins
try {
    $stats.DomainAdmins = (Get-ADGroupMember -Identity "Domain Admins" -ErrorAction SilentlyContinue).Count
    Write-Host "Total de administradores de dom�nio: $($stats.DomainAdmins)" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao contar administradores de dom�nio: $_" -ForegroundColor Yellow
}

# Enterprise Admins
try {
    $stats.EnterpriseAdmins = (Get-ADGroupMember -Identity "Enterprise Admins" -ErrorAction SilentlyContinue).Count
    Write-Host "Total de administradores enterprise: $($stats.EnterpriseAdmins)" -ForegroundColor Green
}
catch {
    $stats.EnterpriseAdmins = 0
    Write-Host "Grupo Enterprise Admins n�o encontrado ou erro ao contar" -ForegroundColor Yellow
}

# GPOs
try {
    $stats.TotalGPOs = (Get-GPO -All).Count
    Write-Host "Total de GPOs: $($stats.TotalGPOs)" -ForegroundColor Green
}
catch {
    $stats.TotalGPOs = 0
    Write-Host "Erro ao contar GPOs: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Analisando sistemas operacionais..." -ForegroundColor Cyan

# An�lise de sistemas operacionais
$osList = @{}
try {
    $computers = Get-ADComputer -Filter * -Properties OperatingSystem
    
    foreach ($computer in $computers) {
        if ($computer.OperatingSystem) {
            $os = $computer.OperatingSystem
            # Simplificar nomes para agrupamento
            if ($os -like "*Windows 10*") { $os = "Windows 10" }
            elseif ($os -like "*Windows 11*") { $os = "Windows 11" }
            elseif ($os -like "*Windows Server 2016*") { $os = "Windows Server 2016" }
            elseif ($os -like "*Windows Server 2019*") { $os = "Windows Server 2019" }
            elseif ($os -like "*Windows Server 2022*") { $os = "Windows Server 2022" }
            
            if ($osList.ContainsKey($os)) {
                $osList[$os]++
            } else {
                $osList[$os] = 1
            }
        } else {
            if ($osList.ContainsKey("Desconhecido")) {
                $osList["Desconhecido"]++
            } else {
                $osList["Desconhecido"] = 1
            }
        }
    }
    
    # Exibir resumo de SO
    Write-Host "Distribui��o de sistemas operacionais:" -ForegroundColor Green
    foreach ($key in $osList.Keys) {
        Write-Host "  $key : $($osList[$key])" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Erro ao analisar sistemas operacionais: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Realizando an�lise de seguran�a..." -ForegroundColor Cyan

# An�lise de seguran�a
$securityAnalysis = @{
    RiskLevel = "M�dio"
    Findings = @()
    Recommendations = @()
}

# Verificar porcentagem de contas com senha que nunca expira
if ($stats.TotalUsers -gt 0) {
    $neverExpiresPercentage = [math]::Round(($stats.PasswordNeverExpires / $stats.EnabledUsers) * 100, 1)
    if ($neverExpiresPercentage -gt 10) {
        $securityAnalysis.Findings += "Alto percentual de contas ($neverExpiresPercentage%) com senhas que nunca expiram"
        $securityAnalysis.Recommendations += "Revisar pol�tica de senhas e configurar expira��o para contas n�o cr�ticas"
    }
}

# Verificar porcentagem de contas sem login recente
if ($stats.EnabledUsers -gt 0) {
    $inactivePercentage = [math]::Round(($stats.LastLogon90Days / $stats.EnabledUsers) * 100, 1)
    if ($inactivePercentage -gt 15) {
        $securityAnalysis.Findings += "Alto percentual de contas ativas ($inactivePercentage%) sem login nos �ltimos 90 dias"
        $securityAnalysis.Recommendations += "Revisar e desativar contas inativas para reduzir superf�cie de ataque"
    }
}

# Verificar n�mero de administradores de dom�nio
if ($stats.DomainAdmins -gt 5) {
    $securityAnalysis.Findings += "N�mero elevado de administradores de dom�nio ($($stats.DomainAdmins))"
    $securityAnalysis.Recommendations += "Reduzir o n�mero de contas com privil�gios de administrador de dom�nio"
}

# Verificar se h� Enterprise Admins (se aplic�vel)
if ($stats.EnterpriseAdmins -gt 2) {
    $securityAnalysis.Findings += "N�mero elevado de administradores enterprise ($($stats.EnterpriseAdmins))"
    $securityAnalysis.Recommendations += "Restringir os privil�gios de administrador enterprise ao m�nimo necess�rio"
}

# Verificar n�vel funcional de dom�nio
$outdatedDomainLevel = $false
if ($stats.DomainLevel -like "*2008*" -or $stats.DomainLevel -like "*2003*" -or $stats.DomainLevel -like "*2000*") {
    $outdatedDomainLevel = $true
    $securityAnalysis.Findings += "N�vel funcional de dom�nio desatualizado ($($stats.DomainLevel))"
    $securityAnalysis.Recommendations += "Atualizar o n�vel funcional do dom�nio para uma vers�o mais recente"
}

# Verificar n�mero de controladores de dom�nio
if ($stats.DomainControllers -lt 2) {
    $securityAnalysis.Findings += "Apenas um controlador de dom�nio detectado"
    $securityAnalysis.Recommendations += "Implementar pelo menos dois controladores de dom�nio para redund�ncia"
}

# Definir n�vel de risco geral com base na quantidade de problemas encontrados
if ($securityAnalysis.Findings.Count -ge 3) {
    $securityAnalysis.RiskLevel = "Alto"
} elseif ($securityAnalysis.Findings.Count -ge 1) {
    $securityAnalysis.RiskLevel = "M�dio"
} else {
    $securityAnalysis.RiskLevel = "Baixo"
}

Write-Host "N�vel de risco identificado: $($securityAnalysis.RiskLevel)" -ForegroundColor Cyan
foreach ($finding in $securityAnalysis.Findings) {
    Write-Host "  Problema: $finding" -ForegroundColor Yellow
}
foreach ($recommendation in $securityAnalysis.Recommendations) {
    Write-Host "  Recomenda��o: $recommendation" -ForegroundColor Green
}

Write-Host ""
Write-Host "Coletando dados para relat�rio..." -ForegroundColor Cyan

# Coletar dados para as tabelas
$topUsers = @()
$disabledUsers = @()
$adminUsers = @()
$servers = @()

try {
    $topUsers = Get-ADUser -Filter * -Properties Name, SamAccountName, EmailAddress, Department, Title, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires | 
        Select-Object Name, SamAccountName, EmailAddress, Department, Title, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires |
        Sort-Object -Property LastLogonDate -Descending |
        Select-Object -First 10
}
catch {
    Write-Host "Erro ao coletar dados de usu�rios ativos: $_" -ForegroundColor Yellow
}

try {
    $disabledUsers = Search-ADAccount -AccountDisabled -UsersOnly |
        Get-ADUser -Properties Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate |
        Select-Object Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate |
        Sort-Object -Property LastLogonDate -Descending |
        Select-Object -First 10
}
catch {
    Write-Host "Erro ao coletar dados de usu�rios desativados: $_" -ForegroundColor Yellow
}

try {
    $adminUsers = Get-ADGroupMember -Identity "Domain Admins" |
        Get-ADUser -Properties Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate, PasswordLastSet |
        Select-Object Name, SamAccountName, EmailAddress, Department, Title, LastLogonDate, PasswordLastSet
}
catch {
    Write-Host "Erro ao coletar dados de administradores: $_" -ForegroundColor Yellow
}

try {
    $servers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' } -Properties Name, OperatingSystem, IPv4Address, Description, LastLogonDate |
        Select-Object Name, OperatingSystem, IPv4Address, Description, LastLogonDate |
        Sort-Object -Property LastLogonDate -Descending |
        Select-Object -First 10
}
catch {
    Write-Host "Erro ao coletar dados de servidores: $_" -ForegroundColor Yellow
}

# Preparar dados para o relat�rio
$userDataHtml = ""
foreach ($user in $topUsers) {
    $statusBadge = ""
    if ($user.Enabled) {
        $statusBadge = '<span class="badge-status badge-success">Ativo</span>'
    } else {
        $statusBadge = '<span class="badge-status badge-danger">Desativado</span>'
    }
    
    $passwordBadge = ""
    if ($user.PasswordNeverExpires) {
        $passwordBadge = '<span class="badge-status badge-warning">Sim</span>'
    } else {
        $passwordBadge = '<span class="badge-status badge-success">N�o</span>'
    }
    
    $userDataHtml += @"
                            <tr>
                                <td>$($user.Name)</td>
                                <td>$($user.SamAccountName)</td>
                                <td>$($user.EmailAddress)</td>
                                <td>$($user.Department)</td>
                                <td>$($user.Title)</td>
                                <td>$(if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" })</td>
                                <td>$statusBadge</td>
                                <td>$passwordBadge</td>
                            </tr>
"@
}

$adminDataHtml = ""
foreach ($admin in $adminUsers) {
    $adminDataHtml += @"
                            <tr>
                                <td>$($admin.Name)</td>
                                <td>$($admin.SamAccountName)</td>
                                <td>$($admin.EmailAddress)</td>
                                <td>$($admin.Department)</td>
                                <td>$($admin.Title)</td>
                                <td>$(if ($admin.LastLogonDate) { $admin.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" })</td>
                                <td>$(if ($admin.PasswordLastSet) { $admin.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Desconhecido" })</td>
                            </tr>
"@
}

$serverDataHtml = ""
foreach ($server in $servers) {
    $serverDataHtml += @"
                            <tr>
                                <td>$($server.Name)</td>
                                <td>$($server.OperatingSystem)</td>
                                <td>$($server.IPv4Address)</td>
                                <td>$($server.Description)</td>
                                <td>$(if ($server.LastLogonDate) { $server.LastLogonDate.ToString("yyyy-MM-dd") } else { "Nunca" })</td>
                            </tr>
"@
}

$findingsHtml = ""
foreach ($finding in $securityAnalysis.Findings) {
    $findingsHtml += @"
                        <div class="warning-item warning-yellow">
                            <i class="fas fa-exclamation-triangle"></i>
                            <span>$finding</span>
                        </div>
"@
}

$recommendationsHtml = ""
foreach ($recommendation in $securityAnalysis.Recommendations) {
    $recommendationsHtml += @"
                        <div class="warning-item warning-green">
                            <i class="fas fa-check-circle"></i>
                            <span>$recommendation</span>
                        </div>
"@
}

# Gerar sistema operacional para o gr�fico
$osLabelsJs = "["
$osDataJs = "["
$osColorsJs = "["

$colorPalette = @(
    '#6a3094', '#9657c7', '#8244b2', '#2c1445', '#c9a6e9', 
    '#5d2683', '#4c1d6b', '#e0cdf0', '#b688df', '#7e3fa8'
)

$i = 0
foreach ($os in $osList.Keys) {
    $osLabelsJs += "'$os',"
    $osDataJs += "$($osList[$os]),"
    $osColorsJs += "'$($colorPalette[$i % $colorPalette.Count])',"
    $i++
}

$osLabelsJs = $osLabelsJs.TrimEnd(',') + "]"
$osDataJs = $osDataJs.TrimEnd(',') + "]"
$osColorsJs = $osColorsJs.TrimEnd(',') + "]"

# Gerar conte�do do corpo do relat�rio
$bodyContent = @"
<!-- Cabe�alho padr�o -->
<div class="header">
    <h1>Active Directory - An�lise Completa</h1>
    <div class="header-actions">
        <button onclick="exportToPdf()"><i class="fas fa-file-pdf"></i> Exportar PDF</button>
        <button onclick="window.print()"><i class="fas fa-print"></i> Imprimir</button>
        <button onclick="sendReport()"><i class="fas fa-envelope"></i> Enviar Relat�rio</button>
    </div>
</div>

<!-- Dashboard -->
<div class="row">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Informa��es do Dom�nio</div>
            <div class="card-body">
                <p><strong>Empresa:</strong> $company</p>
                <p><strong>Dom�nio:</strong> $($stats.DomainName)</p>
                <p><strong>N�vel Funcional de Dom�nio:</strong> $($stats.DomainLevel)</p>
                <p><strong>N�vel Funcional de Floresta:</strong> $($stats.ForestLevel)</p>
                <p><strong>Data:</strong> $date</p>
                <p><strong>Respons�vel:</strong> $owner</p>
            </div>
        </div>
    </div>

    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">
                <span>Resumo da Seguran�a</span>
                <span class="risk-badge risk-badge-$($securityAnalysis.RiskLevel.ToLower())">Risco $($securityAnalysis.RiskLevel)</span>
            </div>
            <div class="card-body">
                <div class="info-box">
                    <p><strong>Total de Usu�rios:</strong> $($stats.TotalUsers) (Ativos: $($stats.EnabledUsers), Desativados: $($stats.DisabledUsers))</p>
                    <p><strong>Usu�rios com senha que nunca expira:</strong> $($stats.PasswordNeverExpires)</p>
                    <p><strong>Usu�rios sem login nos �ltimos $($stats.Days) dias:</strong> $($stats.LastLogon90Days)</p>
                    <p><strong>Administradores de Dom�nio:</strong> $($stats.DomainAdmins)</p>
                </div>
                
                <!-- Problemas encontrados -->
$findingsHtml
                
                <!-- Recomenda��es -->
$recommendationsHtml
            </div>
        </div>
    </div>
</div>

<!-- Estat�sticas -->
<div class="row mb-4">
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-users"></i>
            <h3>$($stats.TotalUsers)</h3>
            <p>Total de Usu�rios</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-desktop"></i>
            <h3>$($stats.TotalComputers)</h3>
            <p>Computadores</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-server"></i>
            <h3>$($stats.TotalServers)</h3>
            <p>Servidores</p>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card stat-card">
            <i class="fas fa-shield-alt"></i>
            <h3>$($stats.DomainControllers)</h3>
            <p>Controladores de Dom�nio</p>
        </div>
    </div>
</div>

<!-- Gr�ficos de Resumo -->
<div class="row mb-4">
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribui��o de Usu�rios</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="userChart"></canvas>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-6">
        <div class="card mb-4">
            <div class="card-header">Distribui��o de Sistemas Operacionais</div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="osChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Tabs para navega��o de dados -->
<div class="card">
    <div class="card-header">
        <ul class="nav nav-tabs card-header-tabs" id="dataTabs" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active" id="users-tab" data-bs-toggle="tab" data-bs-target="#users" type="button" role="tab" aria-controls="users" aria-selected="true">Usu�rios Ativos</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="admins-tab" data-bs-toggle="tab" data-bs-target="#admins" type="button" role="tab" aria-controls="admins" aria-selected="false">Administradores</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="servers-tab" data-bs-toggle="tab" data-bs-target="#servers" type="button" role="tab" aria-controls="servers" aria-selected="false">Servidores</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="security-tab" data-bs-toggle="tab" data-bs-target="#security" type="button" role="tab" aria-controls="security" aria-selected="false">Seguran�a</button>
            </li>
        </ul>
    </div>
    <div class="card-body">
        <div class="tab-content" id="dataTabsContent">
            <!-- Tab de Usu�rios Ativos -->
            <div class="tab-pane fade show active" id="users" role="tabpanel" aria-labelledby="users-tab">
                <h4>Usu�rios Ativos Recentes (Top 10)</h4>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Nome</th>
                                <th>Login</th>
                                <th>Email</th>
                                <th>Departamento</th>
                                <th>Cargo</th>
                                <th>�ltimo Login</th>
                                <th>Status</th>
                                <th>Senha Nunca Expira</th>
                            </tr>
                        </thead>
                        <tbody>
$userDataHtml
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Tab de Administradores -->
            <div class="tab-pane fade" id="admins" role="tabpanel" aria-labelledby="admins-tab">
                <h4>Administradores de Dom�nio</h4>
                <div class="table-responsive">
                    <table>
<thead>
                            <tr>
                                <th>Nome</th>
                                <th>Login</th>
                                <th>Email</th>
                                <th>Departamento</th>
                                <th>Cargo</th>
                                <th>�ltimo Login</th>
                                <th>�ltima Troca de Senha</th>
                            </tr>
                        </thead>
                        <tbody>
$adminDataHtml
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Tab de Servidores -->
            <div class="tab-pane fade" id="servers" role="tabpanel" aria-labelledby="servers-tab">
                <h4>Servidores Windows</h4>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Nome</th>
                                <th>Sistema Operacional</th>
                                <th>Endere�o IP</th>
                                <th>Descri��o</th>
                                <th>�ltimo Login</th>
                            </tr>
                        </thead>
                        <tbody>
$serverDataHtml
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Tab de Seguran�a -->
            <div class="tab-pane fade" id="security" role="tabpanel" aria-labelledby="security-tab">
                <h4>An�lise de Seguran�a</h4>
                
                <div class="mb-4">
                    <h5>N�vel de Risco: <span class="risk-$($securityAnalysis.RiskLevel.ToLower())">$($securityAnalysis.RiskLevel)</span></h5>
                    
                    <div class="row">
                        <div class="col-md-6">
                            <div class="card mb-3">
                                <div class="card-header">Problemas Encontrados</div>
                                <div class="card-body">
                                    <ul class="list-group">
$(foreach ($finding in $securityAnalysis.Findings) {
    "                                        <li class='list-group-item'><i class='fas fa-exclamation-triangle text-warning'></i> $finding</li>"
})
                                    </ul>
                                </div>
                            </div>
                        </div>
                        
                        <div class="col-md-6">
                            <div class="card mb-3">
                                <div class="card-header">Recomenda��es</div>
                                <div class="card-body">
                                    <ul class="list-group">
$(foreach ($recommendation in $securityAnalysis.Recommendations) {
    "                                        <li class='list-group-item'><i class='fas fa-check-circle text-success'></i> $recommendation</li>"
})
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="card">
                    <div class="card-header">M�tricas de Seguran�a</div>
                    <div class="card-body">
                        <div class="row">
                            <div class="col-md-4">
                                <div class="stat-card">
                                    <i class="fas fa-key text-warning"></i>
                                    <h3>$($stats.PasswordNeverExpires)</h3>
                                    <p>Senhas Nunca Expiram</p>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="stat-card">
                                    <i class="fas fa-user-clock text-danger"></i>
                                    <h3>$($stats.LastLogon90Days)</h3>
                                    <p>Sem Login (90 dias)</p>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="stat-card">
                                    <i class="fas fa-user-shield text-primary"></i>
                                    <h3>$($stats.DomainAdmins)</h3>
                                    <p>Administradores</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
"@

# Estilos adicionais para as badges de risco
$extraStyles = @"
<style>
    /* Badges de risco */
    .risk-badge {
        display: inline-block;
        padding: 8px 15px;
        border-radius: 30px;
        font-size: 14px;
        font-weight: 600;
        margin-left: 10px;
    }
    
    .risk-badge-baixo {
        background-color: #d4edda;
        color: #155724;
    }
    
    .risk-badge-medio {
        background-color: #fff3cd;
        color: #856404;
    }
    
    .risk-badge-alto {
        background-color: #f8d7da;
        color: #721c24;
    }
    
    /* Cores para texto de risco */
    .risk-baixo {
        color: #28a745;
        font-weight: bold;
    }
    
    .risk-medio {
        color: #ffc107;
        font-weight: bold;
    }
    
    .risk-alto {
        color: #dc3545;
        font-weight: bold;
    }
    
    /* Estilos para as tabs */
    .nav-tabs .nav-link {
        color: var(--lobios-primary);
        border: none;
        padding: 10px 20px;
        border-radius: 0;
        font-weight: 500;
    }
    
    .nav-tabs .nav-link.active {
        color: white;
        background-color: var(--lobios-primary);
        border-bottom: 2px solid var(--lobios-accent);
    }
    
    .nav-tabs .nav-link:hover {
        border-bottom: 2px solid var(--lobios-light);
    }
    
    .tab-content {
        padding: 20px 0;
    }
</style>
"@

# Script espec�fico para esta p�gina
$extraScripts = @"
$extraStyles

<script>
    // Inicializar gr�ficos
    document.addEventListener('DOMContentLoaded', function() {
        // Gr�fico de distribui��o de usu�rios
        const userCtx = document.getElementById('userChart').getContext('2d');
        const userChart = new Chart(userCtx, {
            type: 'pie',
            data: {
                labels: ['Usu�rios Ativos', 'Usu�rios Desativados', 'Senhas Nunca Expiram', 'Sem Login (90 dias)'],
                datasets: [{
                    data: [
                        $($stats.EnabledUsers - $stats.PasswordNeverExpires - $stats.LastLogon90Days), 
                        $($stats.DisabledUsers), 
                        $($stats.PasswordNeverExpires), 
                        $($stats.LastLogon90Days)
                    ],
                    backgroundColor: [
                        '#28a745', // Verde para ativos sem problemas
                        '#dc3545', // Vermelho para desativados
                        '#ffc107', // Amarelo para senhas que nunca expiram
                        '#6c757d'  // Cinza para sem login
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
        
        // Gr�fico de distribui��o de sistemas operacionais
        const osCtx = document.getElementById('osChart').getContext('2d');
        const osChart = new Chart(osCtx, {
            type: 'bar',
            data: {
                labels: $osLabelsJs,
                datasets: [{
                    label: 'Quantidade',
                    data: $osDataJs,
                    backgroundColor: $osColorsJs,
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

    // Fun��es para intera��o
    function exportToPdf() {
        alert('Exportando relat�rio para PDF...');
        // Implementa��o da exporta��o para PDF
    }
    
    function sendReport() {
        alert('Enviando relat�rio por email...');
        // Implementa��o do envio de relat�rio
    }
</script>
"@

# Gerar o HTML completo usando o helper
$html = New-ADRTReport -BodyContent $bodyContent `
                      -Title "Active Directory Report Tool - An�lise Completa" `
                      -ActiveMenu "An�lise Completa" `
                      -CompanyName $company `
                      -DomainName $stats.DomainName `
                      -Date $date `
                      -Owner $owner `
                      -ExtraScripts $extraScripts

# Salvar o HTML no arquivo de sa�da
try {
    # Criar diret�rio se n�o existir
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    [System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($false))
    Write-Host "? Relat�rio HTML gerado com sucesso em: $outputPath" -ForegroundColor Green
}
catch {
    Write-Host "? Erro ao salvar o relat�rio: $_" -ForegroundColor Red
    exit 1
}

# Abrir o relat�rio no navegador
try {
    Start-Process $outputPath
    Write-Host "? Relat�rio aberto no navegador com sucesso" -ForegroundColor Green
}
catch {
    Write-Host "! Erro ao abrir o relat�rio no navegador: $_" -ForegroundColor Yellow
    Write-Host "Voc� pode abrir manualmente o arquivo em: $outputPath" -ForegroundColor Yellow
}

# Fim do script
Write-Host ""
Write-Host "?????????????????????????????????????????????????????????????????" -ForegroundColor Cyan
Write-Host "?                      AN�LISE CONCLU�DA                        ?" -ForegroundColor Cyan
Write-Host "?????????????????????????????????????????????????????????????????" -ForegroundColor Cyan
Write-Host ""
Write-Host "A an�lise completa do Active Directory foi conclu�da com sucesso."
Write-Host "O relat�rio foi gerado e aberto no seu navegador padr�o."
Write-Host ""
Write-Host "Caminho do relat�rio: $outputPath"
Write-Host ""
Write-Host ""
if ($notificationModuleAvailable -and (Get-Variable -Name notificationSent -ErrorAction SilentlyContinue)) {
    if ($notificationSent) {
        Write-Host "Uma notifica��o foi enviada para a equipe de suporte via email e/ou Telegram." -ForegroundColor Green
        Write-Host ""
    }
}
Write-Host "Recomenda��o: Execute esta an�lise periodicamente para monitorar"
Write-Host "a seguran�a e a integridade do seu Active Directory."
Write-Host ""