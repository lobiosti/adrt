<#
.SYNOPSIS
    Executador de relatórios ADRT
.DESCRIPTION
    Script para executar os relatórios do ADRT corrigindo a importação do módulo de templates
.NOTES
    Este script corrige o problema de importação de módulos, funcionando como um "wrapper"
    para os scripts originais de relatório.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$ReportType
)

# Definir o diretório raiz do projeto (onde está o módulo templates)
$rootDir = "C:\Files\ad-assessment\ad-assessment" # Ajuste esse caminho conforme necessário
$modulePath = Join-Path -Path $rootDir -ChildPath "modules\ADRT-Templates.psm1"

# Verificar se o módulo existe
if (-not (Test-Path -Path $modulePath)) {
    Write-Host "ERRO: Módulo de templates não encontrado em: $modulePath" -ForegroundColor Red
    Write-Host "Verifique se o caminho está correto e tente novamente." -ForegroundColor Yellow
    exit 1
}

# Importar o módulo
Import-Module $modulePath -Force -Global
Write-Host "Módulo ADRT-Templates carregado com sucesso!" -ForegroundColor Green

# Mapear os tipos de relatório para os scripts correspondentes
$reportScripts = @{
    "users" = "ad-reports\ad-users\ad-users-modern.ps1"
    "admins" = "ad-reports\ad-admins\ad-admins-modern.ps1"
    "enterprise-admins" = "ad-reports\ad-enterprise-admins\ad-enterprise-admins-modern.ps1"
    "disabled" = "ad-reports\ad-disabled\ad-disabled-modern.ps1"
    "lastlogon" = "ad-reports\ad-lastlogon\ad-lastlogon-modern.ps1"
    "neverexpires" = "ad-reports\ad-neverexpires\ad-neverexpires-modern.ps1"
    "groups" = "ad-reports\ad-groups\ad-groups-modern.ps1"
    "membergroups" = "ad-reports\ad-membergroups\ad-membergroups-modern.ps1"
    "ous" = "ad-reports\ad-ous\ad-ous-modern.ps1"
    "computers" = "ad-reports\ad-computers\ad-computers-modern.ps1"
    "servers" = "ad-reports\ad-servers\ad-servers-modern.ps1"
    "dcs" = "ad-reports\ad-dcs\ad-dcs-modern.ps1"
    "gpos" = "ad-reports\ad-gpos\ad-gpos-modern.ps1"
    "inventory" = "ad-reports\ad-inventory\ad-inventory-modern.ps1"
    "all" = "ad-reports\ad-all\ad-all-modern.ps1"
    "dashboard" = "index-modern.ps1"
}

# Verificar se o tipo de relatório é válido
if (-not $reportScripts.ContainsKey($ReportType)) {
    Write-Host "ERRO: Tipo de relatório '$ReportType' inválido." -ForegroundColor Red
    Write-Host "Relatórios disponíveis:" -ForegroundColor Yellow
    $reportScripts.Keys | ForEach-Object { Write-Host "  - $_" -ForegroundColor Cyan }
    exit 1
}

$scriptPath = $reportScripts[$ReportType]

# Verificar se o script existe
if (-not (Test-Path -Path $scriptPath)) {
    Write-Host "ERRO: Script '$scriptPath' não encontrado." -ForegroundColor Red
    exit 1
}

# Executar o script
Write-Host "Executando relatório: $ReportType ($scriptPath)" -ForegroundColor Green
& $scriptPath

Write-Host "Relatório $ReportType executado com sucesso!" -ForegroundColor Green