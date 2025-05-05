# ADRT-Helper.ps1
# Funções auxiliares para os relatórios ADRT

function New-ADRTHeader {
    param (
        [string]$Title
    )
    
    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lobios - $Title</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --lobios-primary: #6a3094;
            --lobios-secondary: #9657c7;
            --lobios-light: #f7f5fa;
            --lobios-dark: #2c1445;
            --lobios-accent: #8244b2;
            --lobios-danger: #dc3545;
            --lobios-warning: #ffc107;
            --lobios-success: #28a745;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f8f9fa;
            color: #212529;
            margin: 0;
            padding: 0;
        }
        
        .sidebar {
            background-color: var(--lobios-primary);
            color: white;
            height: 100vh;
            position: fixed;
            width: 280px;
            box-shadow: 2px 0 10px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            transition: all 0.3s;
        }
        
        .sidebar-header {
            padding: 20px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }
        
        .sidebar-header img {
            max-width: 180px;
        }
        
        .sidebar-menu {
            padding: 0;
            list-style: none;
            margin-top: 20px;
        }
        
        .sidebar-menu li {
            padding: 12px 20px;
            margin-bottom: 5px;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .sidebar-menu li:hover {
            background-color: var(--lobios-accent);
        }
        
        .sidebar-menu li.active {
            background-color: var(--lobios-secondary);
            border-left: 4px solid white;
        }
        
        .sidebar-menu i {
            margin-right: 10px;
            width: 20px;
            text-align: center;
        }
        
        .main-content {
            margin-left: 280px;
            padding: 20px;
            transition: all 0.3s;
        }
        
        .header {
            background-color: white;
            padding: 15px 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .header h1 {
            font-size: 24px;
            color: var(--lobios-primary);
            margin: 0;
            font-weight: 600;
        }
        
        .header-actions button {
            background-color: var(--lobios-primary);
            color: white;
            border: none;
            padding: 8px 15px;
            border-radius: 5px;
            cursor: pointer;
            transition: all 0.3s;
            font-size: 14px;
            font-weight: 500;
            margin-left: 10px;
        }
        
        .header-actions button:hover {
            background-color: var(--lobios-accent);
        }
        
        .card {
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            margin-bottom: 20px;
            overflow: hidden;
            border: none;
        }
        
        .card-header {
            background-color: var(--lobios-light);
            color: var(--lobios-primary);
            padding: 15px 20px;
            font-weight: 600;
            border-bottom: 1px solid #eee;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .card-body {
            padding: 20px;
        }
        
        .dashboard-section {
            margin-bottom: 30px;
        }
        
        .info-box {
            padding: 15px;
            border-left: 4px solid var(--lobios-primary);
            margin-bottom: 10px;
            background-color: var(--lobios-light);
        }
        
        .warning-item {
            display: flex;
            align-items: center;
            margin-bottom: 10px;
            padding: 10px;
            border-radius: 5px;
        }
        
        .warning-item i {
            font-size: 20px;
            margin-right: 10px;
        }
        
        .warning-yellow {
            background-color: #fff3cd;
            color: #856404;
        }
        
        .warning-red {
            background-color: #f8d7da;
            color: #721c24;
        }
        
        .chart-container {
            position: relative;
            height: 300px;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            margin-top: 20px;
        }
        
        .footer img {
            max-width: 120px;
            margin-bottom: 10px;
        }
        
        .footer p {
            margin: 0;
            color: #6c757d;
            font-size: 14px;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
        }
        
        table th, table td {
            padding: 12px 15px;
            border-bottom: 1px solid #eee;
            text-align: left;
        }
        
        table th {
            background-color: var(--lobios-light);
            color: var(--lobios-primary);
            font-weight: 600;
        }
        
        table tr:hover {
            background-color: #f8f9fa;
        }
        
        .badge-status {
            display: inline-block;
            padding: 5px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }
        
        .badge-success {
            background-color: #d4edda;
            color: #155724;
        }
        
        .badge-danger {
            background-color: #f8d7da;
            color: #721c24;
        }
        
        .badge-warning {
            background-color: #fff3cd;
            color: #856404;
        }
        
        .badge-info {
            background-color: #d1ecf1;
            color: #0c5460;
        }
        
        .badge-recent {
            background-color: #ffdde5;
            color: #d63384;
        }
        
        .badge-old {
            background-color: #a9a9a9;
            color: #343a40;
        }
        
        .action-buttons {
            display: flex;
            gap: 5px;
        }
        
        .action-button {
            background-color: var(--lobios-light);
            color: var(--lobios-primary);
            border: none;
            width: 30px;
            height: 30px;
            border-radius: 5px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .action-button:hover {
            background-color: var(--lobios-primary);
            color: white;
        }
        
        /* Estatísticas de cards */
        .stat-card {
            text-align: center;
            padding: 20px;
            transition: all 0.3s;
        }
        
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.1);
        }
        
        .stat-card i {
            font-size: 32px;
            color: var(--lobios-primary);
            margin-bottom: 15px;
        }
        
        .stat-card h3 {
            font-size: 28px;
            color: var(--lobios-dark);
            margin-bottom: 10px;
            font-weight: 700;
        }
        
        .stat-card p {
            color: #6c757d;
            margin: 0;
            font-size: 16px;
        }
        
        .truncate {
            max-width: 250px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        /* Responsivo */
        @media (max-width: 992px) {
            .sidebar {
                width: 70px;
            }
            
            .sidebar-header img {
                display: none;
            }
            
            .sidebar-menu span {
                display: none;
            }
            
            .sidebar-menu i {
                margin-right: 0;
                font-size: 18px;
            }
            
            .main-content {
                margin-left: 70px;
            }
        }
        
        @media (max-width: 576px) {
            .sidebar {
                display: none;
            }
            
            .main-content {
                margin-left: 0;
            }
        }
    </style>
</head>
<body>
"@
}

function New-ADRTSidebar {
    <#
    .SYNOPSIS
        Gera o menu lateral HTML com categorias colapsáveis
    .DESCRIPTION
        Cria a barra lateral do documento HTML, com menu agrupado por categorias
    .PARAMETER ActiveMenu
        Nome do item de menu que deve aparecer como ativo
    .EXAMPLE
        $sidebar = New-ADRTSidebar -ActiveMenu "Usuários Desativados"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ActiveMenu = ""
    )
    
    # Estrutura de dados para o menu com categorias
    $menuCategories = @{
        "Dashboard" = @{
            Icon = "fas fa-tachometer-alt"
            Items = @(
                @{Name = "Dashboard"; Icon = "fas fa-home"; Url = "../../index-modern.html"}
            )
        }
        "Usuários e Grupos" = @{
            Icon = "fas fa-users"
            Items = @(
                @{Name = "Todos os Usuários"; Icon = "fas fa-users"; Url = "../../ad-reports/ad-users/ad-users-modern.html"}
                @{Name = "Administradores"; Icon = "fas fa-user-shield"; Url = "../../ad-reports/ad-admins/ad-admins-modern.html"}
                @{Name = "Administradores Enterprise"; Icon = "fas fa-user-tie"; Url = "../../ad-reports/ad-enterprise-admins/ad-enterprise-admins-modern.html"}
                @{Name = "Usuários Desativados"; Icon = "fas fa-user-times"; Url = "../../ad-reports/ad-disabled/ad-disabled-modern.html"}
                @{Name = "Último Login"; Icon = "fas fa-clock"; Url = "../../ad-reports/ad-lastlogon/ad-lastlogon-modern.html"}
                @{Name = "Senhas Nunca Expiram"; Icon = "fas fa-key"; Url = "../../ad-reports/ad-neverexpires/ad-neverexpires-modern.html"}
                @{Name = "Todos os Grupos"; Icon = "fas fa-users-cog"; Url = "../../ad-reports/ad-groups/ad-groups-modern.html"}
                @{Name = "Membros de Grupos"; Icon = "fas fa-layer-group"; Url = "../../ad-reports/ad-membergroups/ad-membergroups-modern.html"}
            )
        }
        "Infraestrutura" = @{
            Icon = "fas fa-network-wired"
            Items = @(
                @{Name = "Todas as OUs"; Icon = "fas fa-sitemap"; Url = "../../ad-reports/ad-ous/ad-ous-modern.html"}
                @{Name = "Computadores"; Icon = "fas fa-desktop"; Url = "../../ad-reports/ad-computers/ad-computers-modern.html"}
                @{Name = "Servidores"; Icon = "fas fa-server"; Url = "../../ad-reports/ad-servers/ad-servers-modern.html"}
                @{Name = "Controladores de Domínio"; Icon = "fas fa-shield-alt"; Url = "../../ad-reports/ad-dcs/ad-dcs-modern.html"}
                @{Name = "Todas as GPOs"; Icon = "fas fa-cogs"; Url = "../../ad-reports/ad-gpos/ad-gpos-modern.html"}
                @{Name = "Inventário"; Icon = "fas fa-clipboard-list"; Url = "../../ad-reports/ad-inventory/ad-inventory-modern.html"}
            )
        }
        "Análise" = @{
            Icon = "fas fa-chart-line"
            Items = @(
                @{Name = "Análise Completa"; Icon = "fas fa-chart-bar"; Url = "../../ad-reports/ad-analysis/ad-analysis-modern.html"}
                @{Name = "Análise de Segurança"; Icon = "fas fa-shield-virus"; Url = "../../ad-reports/ad-analysis/ad-analysis-modern.html"}
            )
        }
    }
    
    # Começar HTML da barra lateral
    $sidebarHtml = @"
<!-- Sidebar -->
<div class="sidebar">
    <div class="sidebar-header">
        <img src="../../web/img/lobios-logo.png" alt="Lobios">
    </div>
    <ul class="sidebar-menu">

"@
    
    # Para cada categoria do menu
    foreach ($category in $menuCategories.Keys) {
        $categoryIcon = $menuCategories[$category].Icon
        $categoryId = $category.Replace(" ", "-").ToLower()
        
        # Verificar se algum item da categoria está ativo
        $categoryHasActiveItem = $false
        foreach ($item in $menuCategories[$category].Items) {
            if ($item.Name -eq $ActiveMenu) {
                $categoryHasActiveItem = $true
                break
            }
        }
        
        # Definir a classe de expansão da categoria (expandida se contém o item ativo)
        $categoryExpandedClass = if ($categoryHasActiveItem) { "show" } else { "" }
        $categoryButtonClass = if ($categoryHasActiveItem) { "" } else { "collapsed" }
        
        # Adicionar o cabeçalho da categoria
        $sidebarHtml += @"
        <li>
            <a href="#$categoryId" class="category-header $categoryButtonClass" data-bs-toggle="collapse" aria-expanded="$($categoryHasActiveItem.ToString().ToLower())">
                <i class="$categoryIcon"></i> 
                <span>$category</span>
                <i class="fas fa-chevron-down ms-auto"></i>
            </a>
            <ul class="collapse $categoryExpandedClass" id="$categoryId">

"@
        
        # Adicionar os itens da categoria
        foreach ($item in $menuCategories[$category].Items) {
            $activeClass = if ($item.Name -eq $ActiveMenu) { " active" } else { "" }
            
            $sidebarHtml += @"
                <li class="sub-item$activeClass" data-url="$($item.Url)">
                    <i class="$($item.Icon)"></i> 
                    <span>$($item.Name)</span>
                </li>

"@
        }
        
        # Fechar a lista de itens e a categoria
        $sidebarHtml += @"
            </ul>
        </li>

"@
    }
    
    # Fechar a lista e a sidebar
    $sidebarHtml += @"
    </ul>
</div>

<!-- Adicionar CSS específico para a sidebar colapsável -->
<style>
    .sidebar-menu .category-header {
        display: flex;
        align-items: center;
        text-decoration: none;
        color: white;
        padding: 12px 20px;
        border-left: 4px solid transparent;
    }
    
    .sidebar-menu .category-header:hover {
        background-color: var(--lobios-accent);
    }
    
    .sidebar-menu .fa-chevron-down {
        transition: transform 0.3s;
    }
    
    .sidebar-menu .category-header.collapsed .fa-chevron-down {
        transform: rotate(-90deg);
    }
    
    .sidebar-menu .sub-item {
        padding: 10px 20px 10px 40px;
        cursor: pointer;
        transition: all 0.3s;
        display: flex;
        align-items: center;
    }
    
    .sidebar-menu .sub-item:hover {
        background-color: var(--lobios-accent);
    }
    
    .sidebar-menu .sub-item.active {
        background-color: var(--lobios-secondary);
        border-left: 4px solid white;
    }
    
    .sidebar-menu .sub-item i {
        margin-right: 10px;
        width: 20px;
        text-align: center;
    }
</style>

<!-- Script para manipular os cliques nos itens do menu -->
<script>
    document.addEventListener('DOMContentLoaded', function() {
        // Adicionar navegação para os sub-itens do menu
        document.querySelectorAll('.sidebar-menu .sub-item').forEach(function(item) {
            item.addEventListener('click', function() {
                const url = this.getAttribute('data-url');
                if (url) {
                    window.location.href = url;
                }
            });
        });
    });
</script>
"@
    
    return $sidebarHtml
}

function New-ADRTFooter {
    param (
        [string]$CompanyName,
        [string]$DomainName,
        [string]$Date,
        [string]$Owner,
        [string]$ExtraScripts = ""
    )
    
    return @"
<!-- Footer -->
<div class="footer">
    <img src="../../web/img/lobios-logo-small.png" alt="Lobios">
    <p>ADRT - Active Directory Report Tool v2.0 | Desenvolvido por Lobios Segurança • Tecnologia • Inovação</p>
    <p>Empresa: $CompanyName - Domínio: $DomainName - Data: $Date - Responsável: $Owner</p>
</div>

<!-- Scripts -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>

<script>
    // Função para navegação entre relatórios
    document.addEventListener('DOMContentLoaded', function() {
        // Adicionar navegação para os itens do menu
        document.querySelectorAll('.sidebar-menu li').forEach(function(item) {
            item.addEventListener('click', function() {
                const menuText = this.querySelector('span').textContent;
                switch(menuText) {
                    case 'Dashboard':
                        window.location.href = '../../index-modern.html';
                        break;
                    case 'Usuários':
                        window.location.href = '../../ad-reports/ad-users/ad-users-modern.html';
                        break;
                    case 'Administradores':
                        window.location.href = '../../ad-reports/ad-admins/ad-admins-modern.html';
                        break;
                    case 'Administradores Enterprise':
                        window.location.href = '../../ad-reports/ad-enterprise-admins/ad-enterprise-admins-modern.html';
                        break;
                    case 'Usuários Desativados':
                        window.location.href = '../../ad-reports/ad-disabled/ad-disabled-modern.html';
                        break;
                    case 'Último Login':
                        window.location.href = '../../ad-reports/ad-lastlogon/ad-lastlogon-modern.html';
                        break;
                    case 'Senhas Nunca Expiram':
                        window.location.href = '../../ad-reports/ad-neverexpires/ad-neverexpires-modern.html';
                        break;
                    case 'Grupos':
                        window.location.href = '../../ad-reports/ad-groups/ad-groups-modern.html';
                        break;
                    case 'Membros de Grupos':
                        window.location.href = '../../ad-reports/ad-membergroups/ad-membergroups-modern.html';
                        break;
                    case 'OUs':
                        window.location.href = '../../ad-reports/ad-ous/ad-ous-modern.html';
                        break;
                    case 'Computadores':
                        window.location.href = '../../ad-reports/ad-computers/ad-computers-modern.html';
                        break;
                    case 'Servidores':
                        window.location.href = '../../ad-reports/ad-servers/ad-servers-modern.html';
                        break;
                    case 'Controladores de Domínio':
                        window.location.href = '../../ad-reports/ad-dcs/ad-dcs-modern.html';
                        break;
                    case 'GPOs':
                        window.location.href = '../../ad-reports/ad-gpos/ad-gpos-modern.html';
                        break;
                    case 'Inventário':
                        window.location.href = '../../ad-reports/ad-inventory/ad-inventory-modern.html';
                        break;
                }
            });
        });
    });
</script>

$ExtraScripts
"@
}

function New-ADRTReport {
    param (
        [string]$BodyContent,
        [string]$Title,
        [string]$ActiveMenu,
        [string]$CompanyName,
        [string]$DomainName,
        [string]$Date,
        [string]$Owner,
        [string]$ExtraScripts = ""
    )
    
    $header = New-ADRTHeader -Title $Title
    $sidebar = New-ADRTSidebar -ActiveMenu $ActiveMenu
    $footer = New-ADRTFooter -CompanyName $CompanyName -DomainName $DomainName -Date $Date -Owner $Owner -ExtraScripts $ExtraScripts
    
    return @"
$header
$sidebar
<div class="main-content">
$BodyContent
$footer
</div>
</html>
"@
}