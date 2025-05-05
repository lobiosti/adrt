# Script para verificar a estrutura de diretórios

# Exibir o diretório atual
$currentDir = (Get-Location).Path
Write-Host "Diretório atual: $currentDir" -ForegroundColor Cyan

# Listar arquivos e pastas no diretório atual
Write-Host "Conteúdo do diretório atual:" -ForegroundColor Yellow
Get-ChildItem -Path . | ForEach-Object {
    Write-Host "  $($_.Name)" -ForegroundColor Gray
}

# Verificar subdiretórios específicos
$modulesDir = Join-Path -Path $currentDir -ChildPath "modules"
if (Test-Path -Path $modulesDir) {
    Write-Host "`nPasta 'modules' encontrada. Conteúdo:" -ForegroundColor Green
    Get-ChildItem -Path $modulesDir | ForEach-Object {
        Write-Host "  $($_.Name)" -ForegroundColor Gray
    }
} else {
    Write-Host "`nPasta 'modules' NÃO encontrada no diretório atual" -ForegroundColor Red
}

# Verificar caminho do script atual
$scriptPath = $MyInvocation.MyCommand.Path
Write-Host "`nCaminho do script atual: $scriptPath" -ForegroundColor Cyan

$scriptDir = Split-Path -Parent $scriptPath
Write-Host "Diretório do script: $scriptDir" -ForegroundColor Cyan

# Verificar o diretório pai
$parentDir = Split-Path -Parent $scriptDir
Write-Host "Diretório pai: $parentDir" -ForegroundColor Cyan

# Verificar existência do módulo usando vários métodos
$paths = @(
    (Join-Path -Path $currentDir -ChildPath "modules\ADRT-Templates.psm1"),
    (Join-Path -Path $parentDir -ChildPath "modules\ADRT-Templates.psm1"),
    (Join-Path -Path $scriptDir -ChildPath "modules\ADRT-Templates.psm1"),
    (Join-Path -Path $currentDir -ChildPath "ad-assessment\modules\ADRT-Templates.psm1"),
    "C:\Files\ad-assessment\ad-assessment\modules\ADRT-Templates.psm1"
)

Write-Host "`nVerificando possíveis localizações do módulo ADRT-Templates.psm1:" -ForegroundColor Yellow
foreach ($path in $paths) {
    $exists = Test-Path -Path $path
    $status = if ($exists) { "ENCONTRADO" } else { "NÃO ENCONTRADO" }
    $color = if ($exists) { "Green" } else { "Red" }
    Write-Host "  $status : $path" -ForegroundColor $color
}

Write-Host "`nDigite qualquer tecla para fechar..." -ForegroundColor Cyan
$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null