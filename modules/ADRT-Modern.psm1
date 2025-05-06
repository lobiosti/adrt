# ADRT-Modern.psm1
# Módulo para gerar relatórios modernos para o Active Directory Report Tool (ADRT)
# Autor: Lobios Segurança • Tecnologia • Inovação

function New-ADRTModernReport {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Title,
        
        [Parameter(Mandatory=$true)]
        [object]$Results,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Users", "DisabledUsers", "LastLogon", "PasswordNeverExpires", "Groups", "MemberGroups", "OUs", "Computers", "Servers", "DomainControllers", "GPOs", "Inventory", "DomainAdmins", "EnterpriseAdmins", "All")]
        [string]$Type,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [string]$CompanyName = "Lobios",
        
        [Parameter(Mandatory=$false)]
        [string]$DomainName = (Get-ADDomain).Forest,
        
        [Parameter(Mandatory=$false)]
        [string]$Owner = "Administrador de TI",
        
        [Parameter(Mandatory=$false)]
        [string]$TemplatePath = "$PSScriptRoot\..\templates\modern-template.html",
        
        [Parameter(Mandatory=$false)]
        [string]$ResourcesPath = "$PSScriptRoot\..\web\img"
    )
    
    begin {
        Write-Verbose "Iniciando geração de relatório: $Title"
        
        # Verificar se o template existe
        if (-not (Test-Path -Path $TemplatePath)) {
            # Tentar encontrar o template em um caminho relativo
            $currentDir = (Get-Item -Path ".").FullName
            $altTemplatePath = Join-Path -Path $currentDir -ChildPath "templates\modern-template.html"
            
            if (Test-Path -Path $altTemplatePath) {
                $TemplatePath = $altTemplatePath
                Write-Verbose "Template encontrado em caminho alternativo: $TemplatePath"
            }
            else {
                Write-Error "Template HTML não encontrado em: $TemplatePath"
                return
            }
        }
        
        # Verificar se o diretório de saída existe
        $outputDir = Split-Path -Parent $OutputPath
        if (-not (Test-Path -Path $outputDir)) {
            if ($PSCmdlet.ShouldProcess($outputDir, "Criar diretório")) {
                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
                Write-Verbose "Diretório de saída criado: $outputDir"
            }
        }
        
        # Carregar template HTML
        try {
            $template = Get-Content -Path $TemplatePath -Raw -ErrorAction Stop
            Write-Verbose "Template HTML carregado com sucesso"
        }
        catch {
            Write-Error "Erro ao carregar o template HTML: $_"
            return
        }
        
        # Inicializar estatísticas do AD
        $stats = @{
            TotalUsers = 0
            DisabledUsers = 0
            PasswordNeverExpires = 0
            LastLogon90Days = 0
            TotalComputers = 0
            TotalServers = 0
            TotalGroups = 0
            TotalOUs = 0
            DomainAdmins = 0
            EnterpriseAdmins = 0
            DomainControllers = 0
            TotalGPOs = 0
        }
        
        # Obter estatísticas apenas se não for um relatório "All"
        if ($Type -ne "All") {
            try {
                Write-Verbose "Coletando estatísticas do Active Directory..."
                
                # Usuários
                $stats.TotalUsers = (Get-ADUser -Filter *).Count
                $stats.DisabledUsers = (Search-ADAccount -AccountDisabled -UsersOnly).Count
                
                # Senhas e login
                $days = 90
                $timestamp = (Get-Date).AddDays(-($days))
                $stats.LastLogon90Days = (Get-ADUser -Filter {LastLogonTimeStamp -lt $timestamp -and enabled -eq $true} -Properties LastLogonTimeStamp).Count
                $stats.PasswordNeverExpires = (Get-ADUser -filter * -properties PasswordNeverExpires | Where-Object { $_.PasswordNeverExpires -eq "true" } | Where-Object {$_.enabled -eq "true"}).Count
                
                # Computadores e servidores
                $stats.TotalComputers = (Get-ADComputer -Filter { OperatingSystem -NotLike '*Windows Server*' }).Count
                $stats.TotalServers = (Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' }).Count
                
                # Grupos e OUs
                $stats.TotalGroups = (Get-ADGroup -Filter {name -like "*"}).Count
                $stats.TotalOUs = (Get-ADOrganizationalUnit -Filter {name -like "*"}).Count
                
                # Controladores de domínio
                $stats.DomainControllers = (Get-ADDomainController -Filter *).Count
                
                # Administradores
                $stats.DomainAdmins = (Get-ADGroupMember -Identity "Domain Admins").Count
                
                # Enterprise Admins (apenas no domínio raiz)
                try {
                    $stats.EnterpriseAdmins = (Get-ADGroupMember -Identity "Enterprise Admins").Count
                }
                catch {
                    $stats.EnterpriseAdmins = 0 # Não existe neste domínio
                }
                
                # GPOs
                try {
                    $stats.TotalGPOs = (Get-GPO -All).Count
                }
                catch {
                    $stats.TotalGPOs = 0 # GPO não disponível
                }
                
                Write-Verbose "Estatísticas coletadas com sucesso"
            }
            catch {
                Write-Warning "Erro ao coletar estatísticas: $_"
            }
        }
    }
    
    process {
        # Data atual
        $currentDate = Get-Date -Format "yyyy-MM-dd"
        
        # Preparar dados para o relatório
        $reportData = @{
            title = $Title
            company = $CompanyName
            domain = $DomainName
            date = $currentDate
            owner = $Owner
            totalUsers = $stats.TotalUsers
            disabledUsers = $stats.DisabledUsers
            totalComputers = $stats.TotalComputers
            totalServers = $stats.TotalServers
            passwordNeverExpires = $stats.PasswordNeverExpires
            lastLogon90Days = $stats.LastLogon90Days
            totalGroups = $stats.TotalGroups
            totalOUs = $stats.TotalOUs
            domainControllers = $stats.DomainControllers
            domainAdmins = $stats.DomainAdmins
            enterpriseAdmins = $stats.EnterpriseAdmins
            totalGPOs = $stats.TotalGPOs
            tableData = @()
            osSummary = @{}
            securitySummary = @{}
        }
        
        # Coletar dados de sistema operacional para gráficos
        if ($Type -in @("Computers", "Servers", "DomainControllers", "Inventory", "All")) {
            try {
                $osSummary = @{}
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
                        
                        if ($osSummary.ContainsKey($os)) {
                            $osSummary[$os]++
                        } else {
                            $osSummary[$os] = 1
                        }
                    } else {
                        if ($osSummary.ContainsKey("Desconhecido")) {
                            $osSummary["Desconhecido"]++
                        } else {
                            $osSummary["Desconhecido"] = 1
                        }
                    }
                }
                
                $reportData.osSummary = $osSummary
                Write-Verbose "Dados de SO coletados: $($osSummary.Count) sistemas operacionais diferentes"
            } catch {
                Write-Warning "Erro ao coletar dados de SO: $_"
            }
        }
        
        # Criar resumo de segurança para gráficos
        $reportData.securitySummary = @{
            "Usuários Ativos" = $stats.TotalUsers - $stats.DisabledUsers
            "Usuários Desativados" = $stats.DisabledUsers
            "Senhas Nunca Expiram" = $stats.PasswordNeverExpires
            "Último Login > 90 dias" = $stats.LastLogon90Days
        }
        
        # Processar dados com base no tipo de relatório
        switch ($Type) {
            "Users" {
                Write-Verbose "Processando dados de usuários..."
                $reportData.tableData = $Results | Select-Object Name, SamAccountName, Department, Title, @{Name="Enabled"; Expression={$_.Enabled}}
                $reportData.activeMenu = "Usuários"
            }
            "DisabledUsers" {
                Write-Verbose "Processando dados de usuários desativados..."
                $reportData.tableData = $Results | Select-Object Name, SamAccountName
                $reportData.activeMenu = "Usuários Desativados"
            }
            "LastLogon" {
                Write-Verbose "Processando dados de último login..."
                $reportData.tableData = $Results | Select-Object Name, SamAccountName, @{Name="LastLogon"; Expression={$_.LastLogon}}
                $reportData.activeMenu = "Último Login"
            }
            "PasswordNeverExpires" {
                Write-Verbose "Processando dados de senhas que nunca expiram..."
                $reportData.tableData = $Results | Select-Object Name, SamAccountName
                $reportData.activeMenu = "Senhas Nunca Expiram"
            }
            "Groups" {
                Write-Verbose "Processando dados de grupos..."
                $reportData.tableData = $Results | Select-Object Name, Description
                $reportData.activeMenu = "Grupos"
            }
            "MemberGroups" {
                Write-Verbose "Processando dados de membros de grupos..."
                $reportData.tableData = $Results | Select-Object Name, MemberOf, Members
                $reportData.activeMenu = "Membros de Grupos"
            }
            "OUs" {
                Write-Verbose "Processando dados de OUs..."
                $reportData.tableData = $Results | Select-Object Name, Description, DistinguishedName
                $reportData.activeMenu = "OUs"
            }
            "Computers" {
                Write-Verbose "Processando dados de computadores..."
                $reportData.tableData = $Results | Select-Object Name, OperatingSystem, Description
                $reportData.activeMenu = "Computadores"
            }
            "Servers" {
                Write-Verbose "Processando dados de servidores..."
                $reportData.tableData = $Results | Select-Object Name, OperatingSystem, Description
                $reportData.activeMenu = "Servidores"
            }
            "DomainControllers" {
                Write-Verbose "Processando dados de controladores de domínio..."
                $reportData.tableData = $Results | Select-Object Name, Site, IPv4Address, OperatingSystem
                $reportData.activeMenu = "Controladores de Domínio"
            }
            "GPOs" {
                Write-Verbose "Processando dados de GPOs..."
                $reportData.tableData = $Results | Select-Object DisplayName, Owner, CreationTime, ModificationTime
                $reportData.activeMenu = "GPOs"
            }
            "Inventory" {
                Write-Verbose "Processando dados de inventário..."
                $reportData.tableData = $Results | Select-Object Name, IPv4Address, LastLogonDate, OperatingSystem, Description
                $reportData.activeMenu = "Inventário"
            }
            "DomainAdmins" {
                Write-Verbose "Processando dados de administradores de domínio..."
                $reportData.tableData = $Results | Select-Object Name, SamAccountName
                $reportData.activeMenu = "Administradores de Domínio"
            }
            "EnterpriseAdmins" {
                Write-Verbose "Processando dados de administradores enterprise..."
                $reportData.tableData = $Results | Select-Object Name, SamAccountName
                $reportData.activeMenu = "Administradores Enterprise"
            }
            "All" {
                Write-Verbose "Processando dados gerais..."
                $reportData.tableData = @()
                $reportData.activeMenu = "Dashboard"
            }
        }
        
        # Converter dados para JSON
        Write-Verbose "Convertendo dados para JSON..."
        $jsonData = $reportData | ConvertTo-Json -Depth 5
        
        # Atualizar o template com os dados
        Write-Verbose "Atualizando template com os dados..."
        
        # 1. Substituir o título da página
        $template = $template -replace '<title>.*?</title>', "<title>Lobios - $Title</title>"
        
        # 2. Inserir os dados do relatório no JavaScript
        # Encontrar o bloco de dados de exemplo e substituí-lo
        $exampleDataPattern = '(?s)const exampleData = \{.*?\};'
        $newDataBlock = "const reportData = $jsonData;"
        $template = $template -replace $exampleDataPattern, $newDataBlock
        
        # 3. Atualizar a função de inicialização
        $template = $template -replace 'fillReportData\(exampleData\);', 'fillReportData(reportData);'
        
        # 4. Atualizar menu ativo
        if ($reportData.activeMenu) {
            # Primeiro resetar o menu ativo atual
            $template = $template -replace '<li class="active">', '<li>'
            
            # Depois definir o novo menu ativo
            $menuPattern = switch ($reportData.activeMenu) {
                "Usuários" { '<li><i class="fas fa-users"></i> <span>Usuários</span></li>' }
                "Administradores de Domínio" { '<li><i class="fas fa-user-shield"></i> <span>Administradores</span></li>' }
                "Administradores Enterprise" { '<li><i class="fas fa-user-shield"></i> <span>Administradores</span></li>' }
                "Usuários Desativados" { '<li><i class="fas fa-user-times"></i> <span>Usuários Desativados</span></li>' }
                "Último Login" { '<li><i class="fas fa-clock"></i> <span>Último Login</span></li>' }
                "Senhas Nunca Expiram" { '<li><i class="fas fa-key"></i> <span>Senhas Nunca Expiram</span></li>' }
                "Grupos" { '<li><i class="fas fa-users-cog"></i> <span>Grupos</span></li>' }
                "Membros de Grupos" { '<li><i class="fas fa-users-cog"></i> <span>Grupos</span></li>' }
                "OUs" { '<li><i class="fas fa-sitemap"></i> <span>OUs</span></li>' }
                "Computadores" { '<li><i class="fas fa-desktop"></i> <span>Computadores</span></li>' }
                "Servidores" { '<li><i class="fas fa-server"></i> <span>Servidores</span></li>' }
                "Controladores de Domínio" { '<li><i class="fas fa-shield-alt"></i> <span>Controladores de Domínio</span></li>' }
                "GPOs" { '<li><i class="fas fa-cogs"></i> <span>GPOs</span></li>' }
                "Inventário" { '<li><i class="fas fa-clipboard-list"></i> <span>Inventário</span></li>' }
                "Dashboard" { '<li><i class="fas fa-tachometer-alt"></i> <span>Dashboard</span></li>' }
                default { $null }
            }
            
            if ($menuPattern) {
                $menuReplacement = $menuPattern -replace '<li>', '<li class="active">'
                $template = $template -replace [regex]::Escape($menuPattern), $menuReplacement
            }
        }
        
        # 5. Atualizar caminhos para recursos/imagens
        if ($ResourcesPath -and (Test-Path -Path $ResourcesPath)) {
            $resourcesRelPath = (Resolve-Path -Path $ResourcesPath -Relative).ToString().Replace(".\", "")
            $template = $template -replace 'path/to/lobios-logo.png', "$resourcesRelPath/lobios-logo.png"
            $template = $template -replace 'path/to/lobios-logo-small.png', "$resourcesRelPath/lobios-logo-small.png"
        }
        
        # Salvar o template atualizado no arquivo HTML final
        try {
            if ($PSCmdlet.ShouldProcess($OutputPath, "Gerar relatório HTML")) {
                $template | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
                #$template | Out-File -FilePath $OutputPath -Encoding UTF8NoBOM -Force
                Write-Verbose "Relatório HTML salvo em: $OutputPath"
            }
        }
        catch {
            Write-Error "Erro ao salvar o relatório: $_"
            return
        }
    }
    
    end {
        if (Test-Path -Path $OutputPath) {
            Write-Output "Relatório gerado com sucesso: $OutputPath"
            
            # Abrir relatório no navegador padrão (opcional)
            if ($PSCmdlet.ShouldProcess($OutputPath, "Abrir no navegador padrão")) {
                try {
                    Start-Process $OutputPath
                }
                catch {
                    Write-Warning "Não foi possível abrir o relatório no navegador: $_"
                }
            }
        }
        else {
            Write-Error "Falha ao gerar o relatório em: $OutputPath"
        }
    }
}

function Convert-ADRTScript {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$ScriptPath,
        
        [Parameter(Mandatory=$false, Position=1)]
        [string]$OutputPath = ""
    )
    
    begin {
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Error "O script especificado não existe: $ScriptPath"
            return
        }
        
        if ([string]::IsNullOrEmpty($OutputPath)) {
            $directory = Split-Path -Parent $ScriptPath
            $fileName = Split-Path -Leaf $ScriptPath
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $extension = [System.IO.Path]::GetExtension($fileName)
            $OutputPath = Join-Path -Path $directory -ChildPath "$baseName-modern$extension"
        }
    }
    
    process {
        # Ler o conteúdo do script original
        $scriptContent = Get-Content -Path $ScriptPath -Raw
        
        # Identificar o tipo de script baseado no nome do arquivo ou conteúdo
        $scriptType = "Unknown"
        $scriptName = Split-Path -Leaf $ScriptPath
        
        switch -Wildcard ($scriptName) {
            "ad-users*" { $scriptType = "Users" }
            "ad-disabled*" { $scriptType = "DisabledUsers" }
            "ad-lastlogon*" { $scriptType = "LastLogon" }
            "ad-neverexpires*" { $scriptType = "PasswordNeverExpires" }
            "ad-groups*" { $scriptType = "Groups" }
            "ad-membergroups*" { $scriptType = "MemberGroups" }
            "ad-ous*" { $scriptType = "OUs" }
            "ad-computers*" { $scriptType = "Computers" }
            "ad-servers*" { $scriptType = "Servers" }
            "ad-dcs*" { $scriptType = "DomainControllers" }
            "ad-gpos*" { $scriptType = "GPOs" }
            "ad-inventory*" { $scriptType = "Inventory" }
            "ad-admins*" { $scriptType = "DomainAdmins" }
            "ad-enterprise-admins*" { $scriptType = "EnterpriseAdmins" }
            "ad-all*" { $scriptType = "All" }
            default { $scriptType = "Unknown" }
        }
        
        # Se não conseguiu identificar pelo nome, tenta pelo conteúdo
        if ($scriptType -eq "Unknown") {
            if ($scriptContent -match "Domain Admins") { $scriptType = "DomainAdmins" }
            elseif ($scriptContent -match "Enterprise Admins") { $scriptType = "EnterpriseAdmins" }
            elseif ($scriptContent -match "Disabled Users") { $scriptType = "DisabledUsers" }
            elseif ($scriptContent -match "Last Logon") { $scriptType = "LastLogon" }
            elseif ($scriptContent -match "Password Never Expires") { $scriptType = "PasswordNeverExpires" }
            elseif ($scriptContent -match "All Users") { $scriptType = "Users" }
            elseif ($scriptContent -match "All Groups") { $scriptType = "Groups" }
            elseif ($scriptContent -match "Member Groups") { $scriptType = "MemberGroups" }
            elseif ($scriptContent -match "All OUs") { $scriptType = "OUs" }
            elseif ($scriptContent -match "All Computers") { $scriptType = "Computers" }
            elseif ($scriptContent -match "All Servers") { $scriptType = "Servers" }
            elseif ($scriptContent -match "Domain Controllers") { $scriptType = "DomainControllers" }
            elseif ($scriptContent -match "All GPOs") { $scriptType = "GPOs" }
            elseif ($scriptContent -match "Inventory") { $scriptType = "Inventory" }
            else { $scriptType = "Users" } # Padrão se não conseguir identificar
        }
        
        # Título do relatório (extraído do script original ou definido automaticamente)
        $title = ""
        if ($scriptContent -match "<b>(.*?)<\/b>") {
            $title = $matches[1] -replace "<font.*?>|<\/font>", ""
        }
        else {
            $title = switch ($scriptType) {
                "Users" { "Active Directory - Todos os Usuários" }
                "DisabledUsers" { "Active Directory - Usuários Desativados" }
                "LastLogon" { "Active Directory - Último Login" }
                "PasswordNeverExpires" { "Active Directory - Senhas Nunca Expiram" }
                "Groups" { "Active Directory - Grupos" }
                "MemberGroups" { "Active Directory - Membros de Grupos" }
                "OUs" { "Active Directory - OUs" }
                "Computers" { "Active Directory - Computadores" }
                "Servers" { "Active Directory - Servidores" }
                "DomainControllers" { "Active Directory - Controladores de Domínio" }
                "GPOs" { "Active Directory - GPOs" }
                "Inventory" { "Active Directory - Inventário" }
                "DomainAdmins" { "Active Directory - Administradores de Domínio" }
                "EnterpriseAdmins" { "Active Directory - Administradores Enterprise" }
                "All" { "Active Directory - Dashboard Geral" }
                default { "Active Directory - Relatório" }
            }
        }
        
        # Determinar qual será o caminho de saída padrão baseado no tipo de relatório
        $defaultOutputPath = ""
        if ($scriptContent -match '\$path\\.*?\.html') {
            $defaultOutputPath = $matches[0]
        }
        else {
            $defaultOutputPath = "ad-reports\$scriptType\$scriptType-modern.html"
        }
        
        # Criar o novo script
        $modernScriptContent = @"
<#
.SYNOPSIS
    $title - Versão Moderna
.DESCRIPTION
    Script ADRT modernizado para $title com interface Lobios
.NOTES
    Original: $scriptName
    Convertido por: ADRT-Modern Module
#>

# Importar o módulo ADRT-Modern
`$modulePath = Join-Path -Path `$PSScriptRoot -ChildPath "modules\ADRT-Modern.psm1"
Import-Module `$modulePath -ErrorAction Stop

# Variáveis do script
`$date = Get-Date -Format "yyyy-MM-dd"
`$directoryPath = (Get-Item -Path ".").FullName
`$outputPath = "$defaultOutputPath"

# Criar diretório se não existir
`$outputDir = Split-Path -Path `$outputPath -Parent
if (-not (Test-Path -Path `$outputDir)) {
    New-Item -ItemType Directory -Path `$outputDir -Force | Out-Null
}

# Obter informações de configuração
if (Test-Path -Path "config\config.txt") {
    `$config = Get-Content -Path "config\config.txt"
    `$company = `$config[7]
    `$owner = `$config[9]
}
else {
    `$company = "Lobios"
    `$owner = "Administrador"
}

# Importar módulo ActiveDirectory
Import-Module ActiveDirectory

# Coletar dados (baseado no script original)
Write-Host "Coletando dados do Active Directory..."

"@

        # Extrair a parte de coleta de dados do script original
        $dataCollection = ""
        
        # Identifica as linhas que contêm a coleta de dados
        if ($scriptContent -match "#-- Filter(.*?)#-- Order by") {
            $dataCollection = $matches[1].Trim()
        }
        elseif ($scriptContent -match "#-- Filter(.*?)#-- Display result") {
            $dataCollection = $matches[1].Trim()
        }
        else {
            # Tentar identificar por padrões comuns se não encontrou pelos padrões anteriores
            switch ($scriptType) {
                "Users" {
                    $dataCollection = @'
$users = @(Get-ADUser -filter * -Properties Company, SamAccountName, Name, Mail, Department, Title, PasswordNeverExpires, Enabled, Created, Modified, Info)
$results = @($users | Select-Object Company, SamAccountName, Name, Mail, Department, Title, PasswordNeverExpires, Enabled, Created, Modified, Info)
'@
                }
                "DisabledUsers" {
                    $dataCollection = @'
$disabled = @(Search-ADAccount -AccountDisabled -UsersOnly)
$results = @($disabled | Select-Object SamAccountName, Name)
'@
                }
                "LastLogon" {
                    $dataCollection = @'
$days = 90
$timestamp = (Get-Date).Adddays(-($days))
$lastlogon = @(Get-ADUser -Filter {LastLogonTimeStamp -lt $timestamp -and enabled -eq $true} -Properties *)
$results = @($lastlogon | select-object Name, SamAccountName, @{Label="LastLogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp).ToString('yyyy/MM/dd hh:mm:ss')}}, Info)
'@
                }
                "PasswordNeverExpires" {
                    $dataCollection = @'
$neverexpires = @(Get-ADUser -filter * -properties PasswordNeverExpires | where { $_.PasswordNeverExpires -eq "true" } | where {$_.enabled -eq "true"} )
$results = @($neverexpires | Select-Object Name, SamAccountName)
'@
                }
                "Groups" {
                    $dataCollection = @'
$groups = @(Get-ADGroup -Filter {name -like "*"} -Properties Description | Select Name, Description)
$results = @($groups | Select-Object Name, Description)
'@
                }
                "DomainAdmins" {
                    $dataCollection = @'
$admins = @(Get-ADGroupMember -Identity "Domain Admins")
$results = @($admins | Select-Object Name, SamAccountName)
'@
                }
                default {
                    # Mensagem em comentário para scripts não reconhecidos
                    $dataCollection = @'
# Não foi possível extrair a coleta de dados do script original.
# Substitua esta seção pela lógica de coleta de dados apropriada para este tipo de relatório.
$results = @()
'@
                }
            }
        }
        
        # Adicionar a coleta de dados ao novo script
        $modernScriptContent += $dataCollection
        
        # Adicionar código de ordenação se disponível no script original
        if ($scriptContent -match "#-- Order by(.*?)#-- Display") {
            $orderBy = $matches[1].Trim()
            $modernScriptContent += "`n`n# Ordenar dados`n$orderBy`n"
        }
        
        # Adicionar a parte de geração do relatório
        $modernScriptContent += @"

# Gerar relatório moderno
Write-Host "Gerando relatório moderno..."
New-ADRTModernReport -Title "$title" `
                     -Results `$results `
                     -Type "$scriptType" `
                     -OutputPath `$outputPath `
                     -CompanyName `$company `
                     -DomainName (Get-ADDomain).Forest `
                     -Owner `$owner `
                     -Verbose

Write-Host ""
Write-Host "Relatório gerado com sucesso em: `$outputPath"
Write-Host ""

# Abrir o relatório no navegador
`$fullOutputPath = Join-Path -Path `$directoryPath -ChildPath `$outputPath
Start-Process `$fullOutputPath
"@
        
        # Salvar o novo script
        if ($PSCmdlet.ShouldProcess($OutputPath, "Criar script modernizado")) {
            $modernScriptContent | Out-File -FilePath $OutputPath -Encoding utf8 -Force
            Write-Output "Script modernizado criado com sucesso: $OutputPath"
        }
    }
    
    end {
        if (Test-Path -Path $OutputPath) {
            return $OutputPath
        }
    }
}

# Exportar funções do módulo
Export-ModuleMember -Function New-ADRTModernReport, Convert-ADRTScript